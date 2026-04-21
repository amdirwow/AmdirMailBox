-- AmdirMailBox - AMDIR Server (WoW 3.3.5)
-- Postal-like logic (stable):
--  - Do NOT rely on "open mail" UI for attachments readiness.
--  - Find attachment slots backward using GetInboxItemLink(i, slot).
--  - After TakeInboxItem/TakeInboxMoney/DeleteInboxItem: WAIT until mailbox changes.
--
-- UI:
--  - One nice button (center top of mail window) + popup menu (EasyMenu) TOGGLE on same button
--  - Simple progress bar below
--  - Confirm ONLY for: "Видалити пусті" and "Видалити все"
--
-- Rules:
--  - COD: always skip (no take, no delete)
--  - TAKE (client): take money + all attachments, DOES NOT delete any mails
--  - DEL EMPTY (client): delete mails with 0 items and 0 money (text doesn't matter), skip COD
--  - REFRESH (server): .mbox refresh
--  - PURGE (server): .mbox purge
--  - STOP (client): stop current run
--
-- Sending QoL:
--  - Keep last recipient name in SendMailNameEditBox after sending, highlighted (select all).
--
-- FIX (this version):
--  - If TakeInboxItem hits UNIQUE/UNIQUE-EQUIPPED (only one allowed) and fails,
--    DO NOT stop the whole run. Skip that attachment slot and continue scanning.
--  - Added maximum diagnostics logs to catch exact error strings & state.

local VMB = {
    frame = nil,
    running = false,
    mode = nil,     -- "TAKE" | "DELALL"
    idx = 0,        -- absolute inbox index: from last to first
    delay = 0.15,
    acc = 0,
    startNum = 0,

    -- attachment scanning state (Postal-style)
    attachIdx = 0,      -- current slot to try (ATTACHMENTS_MAX_RECEIVE..1)
    activeMail = 0,     -- which mail idx attachIdx belongs to

    -- blacklist of attachment slots that failed to take (unique/other errors)
    -- key format: "mailIndex:slot" => true
    badSlots = {},

    pending = false,
    pendingSince = 0,
    pendingTimeout = 4.5,
    pendingSnapshot = nil, -- {idx,num,money,items,cod,action,attachIdx}
    pendingFails = 0,

    -- MAX LOGS: enabled by default in this debug build
    debug = false,

    stats = {
        processed = 0, deleted = 0, takenItems = 0, takenMoney = 0, skippedCOD = 0,
        skippedUnique = 0, skippedOtherErrors = 0
    },
    lastAction = "",
    lastIndex = 0,

    -- UI
    bar = nil,
    actionBtn = nil,
    menuFrame = nil,
    menuOpen = false,

    -- Send mail QoL
    lastRecipient = "",
    deferRestoreRecipient = false,
    deferRestoreRecipientAt = 0,
    deferRestoreAttempts = 0,

    -- hooks
    hookedSend = false,
    hookedRecipientBox = false,

    -- when we intentionally ignore a TAKE_ITEM failure and skip slot,
    -- client may still fire MAIL_FAILED; suppress it for a short time
    suppressMailFailedUntil = 0,
    suppressMailFailedIdx = 0,
}

local UpdateProgress -- forward declare

local function Now()
    return GetTime and GetTime() or 0
end

local function Print(msg)
    DEFAULT_CHAT_FRAME:AddMessage("|cff00ff00[Пошта]|r " .. tostring(msg))
end

local function Dbg(fmt, ...)
    if not VMB.debug then return end
    local msg = fmt
    if select("#", ...) > 0 then
        msg = string.format(fmt, ...)
    end
    DEFAULT_CHAT_FRAME:AddMessage("|cff88ccff[ПоштаDBG]|r " .. msg)
end

local function GetMailInfo(i)
    -- Returns: money, cod, itemCount, hasText, sender, subject
    local _, _, sender, subject, money, cod, _, itemCount, _, _, hasText = GetInboxHeaderInfo(i)
    money = money or 0
    cod = cod or 0
    itemCount = itemCount or 0
    hasText = not not hasText
    sender = sender or ""
    subject = subject or ""
    return money, cod, itemCount, hasText, sender, subject
end

local function ClampIdx()
    local num = GetInboxNumItems() or 0
    if VMB.idx > num then
        Dbg("Clamp idx: %d -> %d", tonumber(VMB.idx or 0), tonumber(num))
        VMB.idx = num
    end
end

local function ClearPending(reason)
    if VMB.pending then
        Dbg("PENDING cleared (%s)", tostring(reason or ""))
        if VMB.pendingSnapshot then
            local s = VMB.pendingSnapshot
            Dbg("PENDING snapshot: idx=%s num=%s money=%s items=%s cod=%s action=%s attachIdx=%s pendingFails=%s",
                tostring(s.idx), tostring(s.num), tostring(s.money), tostring(s.items), tostring(s.cod),
                tostring(s.action), tostring(s.attachIdx), tostring(VMB.pendingFails))
        end
    end
    VMB.pending = false
    VMB.pendingSnapshot = nil
    VMB.pendingSince = 0
end

local function MarkPending(action, snap)
    VMB.pending = true
    VMB.pendingSince = Now()
    VMB.pendingSnapshot = snap
    VMB.lastAction = action or VMB.lastAction
    if snap then
        Dbg("MarkPending: action=%s idx=%s attachIdx=%s money=%s items=%s cod=%s num=%s",
            tostring(action), tostring(snap.idx), tostring(snap.attachIdx), tostring(snap.money),
            tostring(snap.items), tostring(snap.cod), tostring(snap.num))
    end
end

-- ============
-- Send mail QoL (Postal-style, minimal)
--  - Remember last recipient after successful send
--  - Restore it into SendMailNameEditBox and select-all
-- ============

AmdirMailBoxDB = AmdirMailBoxDB or { lastRecipient = "" }

local function Trim(s)
    s = tostring(s or "")
    s = string.gsub(s, "^%s+", "")
    s = string.gsub(s, "%s+$", "")
    return s
end

local function RestoreRecipientSelectAll()
    if not SendMailNameEditBox then return end
    local name = Trim(AmdirMailBoxDB.lastRecipient)
    if name == "" then return end
    SendMailNameEditBox:SetText(name)
    SendMailNameEditBox:HighlightText(0, -1) -- select all
    SendMailNameEditBox:SetCursorPosition(0)
    Dbg("QoL: restored recipient='%s' (select-all)", name)
end

local function EnableSendMailQoL()
    if VMB.hookedSend then return end
    VMB.hookedSend = true

    if hooksecurefunc then
        hooksecurefunc("SendMailFrame_Reset", function()
            RestoreRecipientSelectAll()
        end)

        if type(SendMailFrame_SendMail) == "function" then
            hooksecurefunc("SendMailFrame_SendMail", function()
                if not SendMailNameEditBox then return end
                local name = Trim(SendMailNameEditBox:GetText())
                if name ~= "" then
                    AmdirMailBoxDB.lastRecipient = name
                    Dbg("QoL: saved recipient pre-reset='%s'", name)
                else
                    Dbg("QoL: SendMailFrame_SendMail name empty (not saved)")
                end
            end)
        end
    end

    if SendMailNameEditBox and SendMailNameEditBox.HookScript and not VMB.hookedRecipientBox then
        VMB.hookedRecipientBox = true

        SendMailNameEditBox:HookScript("OnEditFocusGained", function(self)
            if (self:GetText() or "") ~= "" then
                self:HighlightText(0, -1)
                self:SetCursorPosition(0)
            end
        end)

        SendMailNameEditBox:HookScript("OnEditFocusLost", function(self)
            local name = Trim(self:GetText())
            if name ~= "" then
                AmdirMailBoxDB.lastRecipient = name
                Dbg("QoL: saved recipient OnEditFocusLost='%s'", name)
            end
        end)

        SendMailNameEditBox:HookScript("OnEnterPressed", function(self)
            local name = Trim(self:GetText())
            if name ~= "" then
                AmdirMailBoxDB.lastRecipient = name
                Dbg("QoL: saved recipient OnEnterPressed='%s'", name)
            end
        end)
    end
end

UpdateProgress = function()
    if not VMB.bar then return end
    if not VMB.running then
        VMB.bar:Hide()
        return
    end

    local total = tonumber(VMB.startNum or 0) or 0
    if total <= 0 then total = GetInboxNumItems() or 0 end
    if total <= 0 then
        VMB.bar:SetMinMaxValues(0, 1)
        VMB.bar:SetValue(0)
        VMB.bar.text:SetText("")
        VMB.bar:Show()
        return
    end

    local done = total - (tonumber(VMB.idx or 0) or 0)
    if done < 0 then done = 0 end
    if done > total then done = total end

    VMB.bar:SetMinMaxValues(0, total)
    VMB.bar:SetValue(done)
    VMB.bar.text:SetText(string.format("%d / %d", done, total))
    VMB.bar:Show()
end

-- ============
-- Run control
-- ============

local function StopRun(reason)
    if not VMB.running then
        UpdateProgress()
        Print("Немає активного процесу.")
        return
    end

    VMB.running = false
    ClearPending("StopRun")
    UpdateProgress()

    local s = VMB.stats
    local tail = reason and (" (" .. reason .. ")") or ""
    Print(("Зупинено%s. Оброблено: %d | Видалено: %d | Предметів: %d | Грошей: %d | COD: %d | UNIQUE: %d | ERR: %d")
        :format(tail, s.processed, s.deleted, s.takenItems, s.takenMoney, s.skippedCOD, s.skippedUnique, s.skippedOtherErrors))
end

local function FinishRun()
    local s = VMB.stats
    VMB.running = false
    ClearPending("FinishRun")
    UpdateProgress()

    Print(("Готово. Оброблено: %d | Видалено: %d | Предметів: %d | Грошей: %d | COD: %d | UNIQUE: %d | ERR: %d")
        :format(s.processed, s.deleted, s.takenItems, s.takenMoney, s.skippedCOD, s.skippedUnique, s.skippedOtherErrors))
end

-- Wait logic: success = mailbox num changed OR header (money/items) changed
local function PendingPoll()
    if not VMB.pending or not VMB.pendingSnapshot then return false end
    local snap = VMB.pendingSnapshot
    local i = snap.idx
    local num = GetInboxNumItems() or 0

    if num ~= snap.num then
        Dbg("PendingPoll: mailbox num changed %s -> %s (action=%s idx=%s)", tostring(snap.num), tostring(num), tostring(snap.action), tostring(i))
        if snap.action == "DELETE_EMPTY" then
            VMB.stats.deleted = VMB.stats.deleted + 1
        elseif snap.action == "TAKE_MONEY" then
            VMB.stats.takenMoney = VMB.stats.takenMoney + 1
        elseif snap.action == "TAKE_ITEM" then
            VMB.stats.takenItems = VMB.stats.takenItems + 1
        end

        ClearPending("mailbox num changed")
        ClampIdx()
        return true
    end

    if i > num then
        ClearPending("idx > num")
        ClampIdx()
        return true
    end

    local money, cod, items = GetMailInfo(i)

    if cod ~= snap.cod then
        ClearPending("cod changed")
        return true
    end

    if money ~= snap.money or items ~= snap.items then
        Dbg("PendingPoll: header changed idx=%s money %s->%s items %s->%s (action=%s)",
            tostring(i), tostring(snap.money), tostring(money), tostring(snap.items), tostring(items), tostring(snap.action))

        if snap.action == "TAKE_MONEY" and money < snap.money then
            VMB.stats.takenMoney = VMB.stats.takenMoney + 1
        elseif snap.action == "TAKE_ITEM" and items < snap.items then
            VMB.stats.takenItems = VMB.stats.takenItems + 1
        elseif snap.action == "DELETE_EMPTY" then
            VMB.stats.deleted = VMB.stats.deleted + 1
        end

        ClearPending("header changed")
        return true
    end

    local dt = Now() - (VMB.pendingSince or 0)
    if dt >= (VMB.pendingTimeout or 4.5) then
        VMB.pendingFails = (VMB.pendingFails or 0) + 1
        Dbg("PENDING no change (fail #%d) idx=%d action=%s attachIdx=%s",
            tonumber(VMB.pendingFails), tonumber(i), tostring(snap.action), tostring(snap.attachIdx))

        if VMB.pendingFails >= 12 then
            StopRun("не вдалося виконати дію (silent fail)")
            return false
        end

        ClearPending("no change -> retry")
        return true
    end

    return false
end

local function StartRun(mode)
    if VMB.running then
        Print("|cffff0000Процес вже запущено.|r")
        return
    end

    if not InboxFrame or not InboxFrame:IsShown() then
        Print("|cffff0000Спочатку відкрий пошту біля скриньки.|r")
        return
    end

    local num = GetInboxNumItems() or 0
    VMB.startNum = num

    if num <= 0 then
        Print("Немає листів.")
        return
    end

    VMB.running = true
    VMB.mode = mode
    VMB.idx = num
    VMB.acc = 0

    VMB.pending = false
    VMB.pendingSnapshot = nil
    VMB.pendingSince = 0
    VMB.pendingFails = 0

    VMB.attachIdx = 0
    VMB.activeMail = 0
    VMB.badSlots = {}

    VMB.stats = {
        processed = 0, deleted = 0, takenItems = 0, takenMoney = 0, skippedCOD = 0,
        skippedUnique = 0, skippedOtherErrors = 0
    }

    if mode == "TAKE" then
        Print("Розпочато забирання вкладень. Листів: " .. tostring(num) .. ".")
    elseif mode == "DELALL" then
        Print("Розпочато видалення порожніх листів. Листів: " .. tostring(num) .. ".")
    end

    Dbg("StartRun: mode=%s num=%s delay=%s pendingTimeout=%s", tostring(mode), tostring(num), tostring(VMB.delay), tostring(VMB.pendingTimeout))
    UpdateProgress()
end

-- ============
-- Attach scan (Postal-style)
-- ============

local function ResetAttachStateForMail(i)
    VMB.activeMail = i
    VMB.attachIdx = ATTACHMENTS_MAX_RECEIVE or 12
    Dbg("ResetAttachStateForMail: idx=%s attachIdx=%s", tostring(i), tostring(VMB.attachIdx))
end

local function FindNextAttachSlot(i)
    if VMB.activeMail ~= i or VMB.attachIdx <= 0 then
        ResetAttachStateForMail(i)
    end

    while VMB.attachIdx > 0 do
        local slot = VMB.attachIdx
        local key = tostring(i) .. ":" .. tostring(slot)

        if VMB.badSlots and VMB.badSlots[key] then
            Dbg("FindNextAttachSlot: idx=%s slot=%s is blacklisted -> skip", tostring(i), tostring(slot))
            VMB.attachIdx = VMB.attachIdx - 1
        else
            local link = GetInboxItemLink(i, slot)
            if link then
                Dbg("FindNextAttachSlot: idx=%s -> slot=%s (link=%s)", tostring(i), tostring(slot), tostring(link))
                return slot
            end
            VMB.attachIdx = VMB.attachIdx - 1
        end
    end

    Dbg("FindNextAttachSlot: idx=%s -> slot=0 (no valid slots)", tostring(i))
    return 0
end

-- ============
-- Error classification (UNIQUE etc.)
-- ============

local function IsUniqueErrorMessage(msgLower)
    -- covers EN/RU/UA common variants on 3.3.5 realms
    if not msgLower or msgLower == "" then return false end

    -- English (typical)
    if string.find(msgLower, "you can only carry") then return true end
    if string.find(msgLower, "unique") then return true end
    if string.find(msgLower, "unique%-equipped") then return true end
    if string.find(msgLower, "only one") then return true end
    if string.find(msgLower, "too many of that item") then return true end
    if string.find(msgLower, "too many") then return true end

    -- Russian
    if string.find(msgLower, "уник") then return true end
    if string.find(msgLower, "только один") then return true end
    if string.find(msgLower, "можете нести") and string.find(msgLower, "один") then return true end
    if string.find(msgLower, "слишком много таких предметов") then return true end

   

    return false
end

local function IsInventoryFullMessage(msgLower)
    if not msgLower or msgLower == "" then return false end
    return (
        string.find(msgLower, "inventory") or string.find(msgLower, "not enough room") or string.find(msgLower, "full")
        or string.find(msgLower, "нема") or string.find(msgLower, "мест") or string.find(msgLower, "сумк")
    ) and not IsUniqueErrorMessage(msgLower)
end

local function IsBusyMessage(msgLower)
    if not msgLower or msgLower == "" then return false end
    return (
        string.find(msgLower, "busy") or string.find(msgLower, "занят") or string.find(msgLower, "can't") or string.find(msgLower, "не мож")
    ) and not IsUniqueErrorMessage(msgLower)
end

-- ============
-- Core tick
-- ============

local function StepProcess()
    if not VMB.running then return end
    UpdateProgress()

    if not InboxFrame or not InboxFrame:IsShown() then
        StopRun("mailbox closed")
        return
    end

    if VMB.pending then
        PendingPoll()
        return
    end

    ClampIdx()

    local i = VMB.idx
    if not i or i <= 0 then
        FinishRun()
        return
    end

    local num = GetInboxNumItems() or 0
    if i > num then
        VMB.idx = num
        return
    end

    local money, cod, itemCount, hasText, sender, subject = GetMailInfo(i)
    VMB.lastIndex = i

    Dbg("STEP idx=%d/%d money=%d cod=%d items=%d text=%s sender=%s subj=%s activeMail=%s attachIdx=%s",
        i, num, money, cod, itemCount, tostring(hasText), tostring(sender), tostring(subject),
        tostring(VMB.activeMail), tostring(VMB.attachIdx))

    -- COD skip always
    if cod and cod > 0 then
        VMB.stats.skippedCOD = VMB.stats.skippedCOD + 1
        Dbg("SKIP: COD idx=%d cod=%d", i, cod)
        VMB.idx = i - 1
        return
    end

    -- DELALL: delete only truly empty (no items, no money), regardless of text
    if VMB.mode == "DELALL" then
        if itemCount == 0 and money == 0 then
            VMB.lastAction = "DELETE_EMPTY"
            Dbg("ACTION: DeleteInboxItem(%d) (empty, text=%s)", i, tostring(hasText))
            DeleteInboxItem(i)
            MarkPending("DELETE_EMPTY", { idx=i, num=num, money=money, items=itemCount, cod=cod, action="DELETE_EMPTY" })
            return
        end

        VMB.stats.processed = VMB.stats.processed + 1
        VMB.idx = i - 1
        return
    end

    -- TAKE: money first
    if money > 0 then
        VMB.lastAction = "TAKE_MONEY"
        Dbg("ACTION: TakeInboxMoney(%d) money=%d", i, money)
        TakeInboxMoney(i)
        MarkPending("TAKE_MONEY", { idx=i, num=num, money=money, items=itemCount, cod=cod, action="TAKE_MONEY" })
        return
    end

    -- TAKE: attachments (scan slots backwards)
    if itemCount > 0 then
        local slot = FindNextAttachSlot(i)
        if slot <= 0 then
            Dbg("ATTACH scan found no slots (idx=%d items=%d) -> retry", i, itemCount)
            MarkPending("SCAN_WAIT", { idx=i, num=num, money=money, items=itemCount, cod=cod, action="SCAN_WAIT" })
            return
        end

        local link = GetInboxItemLink(i, slot)
        VMB.lastAction = "TAKE_ITEM"
        Dbg("ACTION: TakeInboxItem(%d,%d) items=%d link=%s", i, slot, itemCount, tostring(link))

        TakeInboxItem(i, slot)

        -- Next scan should continue below current slot (unless error forces other behavior)
        VMB.attachIdx = slot - 1

        MarkPending("TAKE_ITEM", { idx=i, num=num, money=money, items=itemCount, cod=cod, action="TAKE_ITEM", attachIdx=slot })
        return
    end

    -- nothing to take
    VMB.stats.processed = VMB.stats.processed + 1
    VMB.idx = i - 1
end

-- ============
-- UI: Action button + EasyMenu (toggle open/close)
-- ============

local MENU = {
    { key="TAKE",        text="Забрати вкладення", confirm=false },
    { key="DELALL",      text="Видалити пусті",    confirm=true  },
    { key="REFRESH_SRV", text="Оновити",           confirm=false },
    { key="PURGE_SRV",   text="Видалити все",      confirm=true  },
    { key="STOP",        text="Стоп",              confirm=false },
}

local function FindMenuEntry(key)
    for _, v in ipairs(MENU) do
        if v.key == key then return v end
    end
    return nil
end

local function DoAction(key)
    if key == "TAKE" then StartRun("TAKE"); return end
    if key == "DELALL" then StartRun("DELALL"); return end
    if key == "REFRESH_SRV" then SendChatMessage(".mbox refresh", "SAY"); return end
    if key == "PURGE_SRV" then SendChatMessage(".mbox purge", "SAY"); return end
    if key == "STOP" then StopRun("manual"); return end
end

local function CloseMenu()
    if CloseDropDownMenus then
        CloseDropDownMenus()
    end
    VMB.menuOpen = false
end

local function AskConfirmAndDo(key)
    local entry = FindMenuEntry(key)
    if not entry then return end

    CloseMenu()

    if not entry.confirm then
        DoAction(key)
        return
    end

    StaticPopupDialogs["VMB_CONFIRM_ONLY_SOME"] = {
        text = entry.text .. "?",
        button1 = "Так",
        button2 = "Нi",
        OnAccept = function() DoAction(key) end,
        timeout = 0,
        whileDead = true,
        hideOnEscape = true,
    }
    StaticPopup_Show("VMB_CONFIRM_ONLY_SOME")
end

local function BuildEasyMenuList()
    local list = {}
    for _, v in ipairs(MENU) do
        list[#list + 1] = {
            text = v.text,
            notCheckable = true,
            func = function() AskConfirmAndDo(v.key) end,
        }
    end
    return list
end

local function ShowActionMenu(anchor)
    if not EasyMenu then
        Print("|cffff0000EasyMenu не знайдено (UI).|r")
        return
    end

    if not VMB.menuFrame then
        VMB.menuFrame = CreateFrame("Frame", "VMB_EasyMenuFrame", UIParent, "UIDropDownMenuTemplate")
        VMB.menuFrame:SetScript("OnHide", function()
            VMB.menuOpen = false
        end)
    end

    local menuList = BuildEasyMenuList()
    EasyMenu(menuList, VMB.menuFrame, anchor, 0, 0, "MENU", 2)
    VMB.menuOpen = true
end

local function CreateUI()
    if VMB.frame then return end
    if not InboxFrame then return end

    -- Minimal holder: NO backdrop (so no "transparent background around button")
    local holder = CreateFrame("Frame", "VMB_Holder", InboxFrame)
    VMB.frame = holder
    holder:SetSize(260, 24)
    holder:SetPoint("TOP", InboxFrame, "TOP", 0, -48)

    -- Action button (always shows same text)
    local btn = CreateFrame("Button", "VMB_ActionButton", holder, "UIPanelButtonTemplate")
    VMB.actionBtn = btn
    btn:SetSize(240, 22)
    btn:SetPoint("CENTER", holder, "CENTER", 0, 0)
    btn:SetText("Оберіть дію")

    btn:SetScript("OnClick", function(self)
        if VMB.menuOpen then
            CloseMenu()
        else
            ShowActionMenu(self)
        end
    end)

    btn:SetScript("OnEnter", function(self)
        GameTooltip:SetOwner(self, "ANCHOR_TOP")
        GameTooltip:SetText("Оберіть дію", 1, 1, 1)
        GameTooltip:Show()
    end)
    btn:SetScript("OnLeave", function() GameTooltip:Hide() end)

    -- Progress bar (below)
    local bar = CreateFrame("StatusBar", "VMB_ProgressBar", InboxFrame)
    VMB.bar = bar
    bar:SetPoint("TOP", holder, "BOTTOM", 0, -2)
    bar:SetSize(280, 10)
    bar:SetStatusBarTexture("Interface\\TargetingFrame\\UI-StatusBar")
    bar:SetMinMaxValues(0, 1)
    bar:SetValue(0)
    bar:Hide()

    bar.bg = bar:CreateTexture(nil, "BACKGROUND")
    bar.bg:SetAllPoints(true)
    bar.bg:SetTexture("Interface\\Buttons\\WHITE8x8")
    bar.bg:SetVertexColor(0, 0, 0, 0.35)

    bar.text = bar:CreateFontString(nil, "OVERLAY", "GameFontNormalSmall")
    bar.text:SetPoint("CENTER", bar, "CENTER", 0, 0)
    bar.text:SetText("")

    -- Ticker
    holder:SetScript("OnUpdate", function(self, elapsed)
        if not VMB.running then return end
        VMB.acc = (VMB.acc or 0) + elapsed
        if VMB.acc >= (VMB.delay or 0.35) then
            VMB.acc = 0
            local ok, err = pcall(StepProcess)
            if not ok then
                Dbg("LUA ERROR in StepProcess: %s", tostring(err))
                StopRun("lua error")
            end
        end
    end)

    EnableSendMailQoL()

    if SendMailFrame and SendMailFrame:IsShown() then
        RestoreRecipientSelectAll()
    end

    Dbg("UI created (debug=%s).", tostring(VMB.debug))
end

-- ============
-- Events
-- ============

local ev = CreateFrame("Frame", nil, UIParent)
ev:RegisterEvent("MAIL_SHOW")
ev:RegisterEvent("MAIL_CLOSED")
ev:RegisterEvent("MAIL_INBOX_UPDATE")
ev:RegisterEvent("MAIL_FAILED")
ev:RegisterEvent("UI_ERROR_MESSAGE")

ev:SetScript("OnEvent", function(self, event, ...)
    if event == "MAIL_SHOW" then
        Dbg("EVT: MAIL_SHOW")
        CreateUI()
        return
    end

    if event == "MAIL_CLOSED" then
        Dbg("EVT: MAIL_CLOSED")
        if VMB.running then StopRun("mailbox closed") end
        CloseMenu()
        return
    end

    if event == "MAIL_INBOX_UPDATE" then
        if VMB.running then
            Dbg("EVT: MAIL_INBOX_UPDATE (idx=%s pending=%s action=%s attachIdx=%s)",
                tostring(VMB.idx), tostring(VMB.pending), tostring(VMB.lastAction), tostring(VMB.attachIdx))
            ClampIdx()
        end
        return
    end

    if event == "MAIL_FAILED" then
        if VMB.running then
            local now = Now() or 0
            local idxNow = tonumber(VMB.lastIndex or VMB.idx or 0) or 0

            Dbg("EVT: MAIL_FAILED (lastAction=%s idx=%d pending=%s suppressIdx=%s suppressUntil=%.2f now=%.2f)",
                tostring(VMB.lastAction), idxNow, tostring(VMB.pending),
                tostring(VMB.suppressMailFailedIdx), tonumber(VMB.suppressMailFailedUntil or 0), now)

            -- If we just handled a TAKE_ITEM fail (unique/other) and intentionally skipped it,
            -- ignore the MAIL_FAILED that often fires right after UI_ERROR_MESSAGE.
            if (now <= (VMB.suppressMailFailedUntil or 0)) and (idxNow == (VMB.suppressMailFailedIdx or -1)) then
                Dbg("MAIL_FAILED suppressed (idx=%d) -> continue run", idxNow)
                return
            end

            if VMB.pendingSnapshot then
                local s = VMB.pendingSnapshot
                Dbg("MAIL_FAILED snapshot: idx=%s num=%s money=%s items=%s cod=%s action=%s attachIdx=%s",
                    tostring(s.idx), tostring(s.num), tostring(s.money), tostring(s.items),
                    tostring(s.cod), tostring(s.action), tostring(s.attachIdx))
            end

            StopRun("MAIL_FAILED")
        end
        return
    end


    if event == "UI_ERROR_MESSAGE" then
        -- Args vary; we log all.
        local a1, a2, a3, a4 = ...
        local msg = tostring(a2 or a1 or a3 or a4 or "")
        local msgLower = string.lower(msg or "")

        Dbg("EVT: UI_ERROR_MESSAGE raw: a1=%s a2=%s a3=%s a4=%s | msg='%s' | lastAction=%s idx=%s pending=%s attachIdx=%s",
            tostring(a1), tostring(a2), tostring(a3), tostring(a4),
            tostring(msg), tostring(VMB.lastAction), tostring(VMB.lastIndex),
            tostring(VMB.pending), tostring(VMB.attachIdx))

        if not VMB.running then return end
        if msg == "" then return end

if IsUniqueErrorMessage(msgLower) then
    VMB.stats.skippedUnique = VMB.stats.skippedUnique + 1

    local snap = VMB.pendingSnapshot
    local curIdx = (snap and snap.idx) or VMB.lastIndex or VMB.idx
    local curSlot = (snap and snap.attachIdx) or 0

    Dbg("UNIQUE detected: idx=%s slot=%s msg='%s' -> BLACKLIST slot, continue.", tostring(curIdx), tostring(curSlot), tostring(msg))

    -- suppress MAIL_FAILED right after this
    VMB.suppressMailFailedIdx = tonumber(curIdx) or 0
    VMB.suppressMailFailedUntil = (Now() or 0) + 0.9

    -- blacklist this slot so we don't retry forever
    if curIdx and curSlot and curSlot > 0 then
        local key = tostring(curIdx) .. ":" .. tostring(curSlot)
        VMB.badSlots[key] = true
        Dbg("BLACKLIST add: %s", key)
    end

    ClearPending("unique -> blacklist")

    -- Force rescan for this mail (but blacklist will prevent повтор)
    VMB.activeMail = curIdx
    VMB.attachIdx = (ATTACHMENTS_MAX_RECEIVE or 12)

    -- If after blacklist there are no other slots, move to next mail to avoid looping
    local nextSlot = FindNextAttachSlot(curIdx)
    if not nextSlot or nextSlot <= 0 then
        Dbg("No more attach slots after UNIQUE blacklist on idx=%s -> next mail", tostring(curIdx))
        VMB.stats.processed = VMB.stats.processed + 1
        VMB.idx = (tonumber(curIdx) or 0) - 1
    end

    VMB.pendingFails = 0
    return
end



        -- Inventory full / not enough room -> STOP (this is correct behavior).
        if IsInventoryFullMessage(msgLower) then
            StopRun("ui_error: " .. msg)
            return
        end

        -- Busy / can't / etc -> STOP (as before), but logged.
        if IsBusyMessage(msgLower) then
            StopRun("ui_error: " .. msg)
            return
        end

        -- Other UI errors:
        -- If it's during TAKE_ITEM, it's safer to SKIP THIS SLOT and continue (instead of killing the run),
        -- but we count it so you can inspect logs and decide.
        if VMB.pending and VMB.pendingSnapshot and VMB.pendingSnapshot.action == "TAKE_ITEM" then
            VMB.stats.skippedOtherErrors = VMB.stats.skippedOtherErrors + 1
            local snap = VMB.pendingSnapshot
            local curIdx = snap.idx
            local curSlot = snap.attachIdx or 0

            Dbg("Non-fatal TAKE_ITEM UI error: idx=%s slot=%s msg='%s' -> BLACKLIST slot, continue.",
                tostring(curIdx), tostring(curSlot), tostring(msg))

            VMB.suppressMailFailedIdx = tonumber(curIdx) or 0
            VMB.suppressMailFailedUntil = (Now() or 0) + 0.9

            if curIdx and curSlot and curSlot > 0 then
                local key = tostring(curIdx) .. ":" .. tostring(curSlot)
                VMB.badSlots[key] = true
                Dbg("BLACKLIST add: %s", key)
            end

            ClearPending("take_item ui_error -> blacklist")

            VMB.activeMail = curIdx
            VMB.attachIdx = (ATTACHMENTS_MAX_RECEIVE or 12)

            local nextSlot = FindNextAttachSlot(curIdx)
            if not nextSlot or nextSlot <= 0 then
                Dbg("No more attach slots after UI-error blacklist on idx=%s -> next mail", tostring(curIdx))
                VMB.stats.processed = VMB.stats.processed + 1
                VMB.idx = (tonumber(curIdx) or 0) - 1
            end

            VMB.pendingFails = 0
            return
        end



        -- Default: keep previous stop behavior for unknown errors outside TAKE_ITEM context.
        StopRun("ui_error: " .. msg)
        return
    end
end)

//
//  main.swift
//  R 角调整 — macOS 菜单栏 app
//
//  行为：
//    1. 在 macOS 顶部菜单栏放一个图标（NSStatusItem）
//    2. 点击图标弹菜单：
//       - 调整 R 角...
//       - 锁定当前选区 R 角
//       - 解锁当前选区
//       - 重新应用锁定
//       ---
//       - 在 Finder 中显示锁定文件
//       - 关于 R 角调整
//       ---
//       - 退出
//    3. 所有操作通过 osascript 调用 PowerPoint AppleScript
//    4. 锁定信息存到 ~/Library/Application Support/RadiusInPpt/locks.json
//

import Cocoa
import os.log

// MARK: - 数据结构

struct ShapeInfo {
    let shapeId: String
    let radiusCm: Double
    let shortSide: Double
}

struct LockEntry: Codable {
    var radiusCm: Double
    var locked: Bool = true
}

typealias LockMap = [String: LockEntry]

// MARK: - 主程序

class AppDelegate: NSObject, NSApplicationDelegate {

    private let log = OSLog(subsystem: "com.jerrrry666.radiusinppt", category: "main")
    private var statusItem: NSStatusItem!
    private let locksFile: URL

    override init() {
        let fm = FileManager.default
        let appSupport = (try? fm.url(
            for: .applicationSupportDirectory,
            in: .userDomainMask,
            appropriateFor: nil,
            create: true
        )) ?? URL(fileURLWithPath: NSHomeDirectory())
        let appDir = appSupport.appendingPathComponent("RadiusInPpt", isDirectory: true)
        try? fm.createDirectory(at: appDir, withIntermediateDirectories: true)
        self.locksFile = appDir.appendingPathComponent("locks.json")
        super.init()
    }

    // MARK: - 启动

    func applicationDidFinishLaunching(_ n: Notification) {
        os_log("R 角调整 菜单栏 app 启动", log: log, type: .info)

        // 1. 菜单栏图标
        statusItem = NSStatusBar.system.statusItem(withLength: NSStatusItem.variableLength)
        if let btn = statusItem.button {
            if let img = loadMenuBarIcon() {
                btn.image = img
            } else {
                btn.title = "R⤢"
            }
            btn.toolTip = "R 角调整 — 点击打开菜单"
        }

        // 2. 菜单
        statusItem.menu = buildMenu()
    }

    private func loadMenuBarIcon() -> NSImage? {
        // 优先从 .app/Contents/Resources/ 加载
        if let path = Bundle.main.path(forResource: "menubar-icon", ofType: "png"),
           let img = NSImage(contentsOfFile: path) {
            img.size = NSSize(width: 18, height: 18)
            img.isTemplate = true  // 跟随菜单栏明暗主题
            return img
        }
        return nil
    }

    private func buildMenu() -> NSMenu {
        let menu = NSMenu()
        menu.autoenablesItems = false

        addMenu(menu, "调整 R 角...", action: #selector(adjustRadius), key: "r")
        addMenu(menu, "锁定当前选区 R 角", action: #selector(lockCurrent), key: "l")
        addMenu(menu, "解锁当前选区", action: #selector(unlockCurrent), key: "")
        addMenu(menu, "重新应用锁定", action: #selector(reapplyLocks), key: "")

        menu.addItem(NSMenuItem.separator())
        addMenu(menu, "在 Finder 中显示锁定文件", action: #selector(revealLocks), key: "")
        addMenu(menu, "关于 R 角调整", action: #selector(showAbout), key: "")

        menu.addItem(NSMenuItem.separator())
        addMenu(menu, "退出", action: #selector(quit), key: "q")

        return menu
    }

    private func addMenu(_ menu: NSMenu, _ title: String, action: Selector, key: String) {
        let item = NSMenuItem(title: title, action: action, keyEquivalent: key)
        item.target = self
        menu.addItem(item)
    }

    // MARK: - 菜单项响应

    @objc func adjustRadius() {
        // 1. 读当前选区
        let info = readSelectionInfo()
        if info.isEmpty {
            showAlert("未选中圆角矩形", "请先在 PowerPoint 中选中一个或多个圆角矩形。")
            return
        }

        // 2. 弹输入框（默认填当前 R 角）
        let defaultValue: String
        if let first = info.first {
            defaultValue = String(format: "%.2f", first.radiusCm)
        } else {
            defaultValue = "0.30"
        }
        guard let input = promptInput(
            title: "调整 R 角",
            message: info.count == 1
                ? "当前选中 1 个圆角矩形（当前 R 角 = \(defaultValue) 厘米）"
                : "当前选中 \(info.count) 个圆角矩形",
            placeholder: "0.30",
            defaultValue: defaultValue
        ) else { return }

        guard let cm = Double(input), cm >= 0 else {
            showAlert("数值无效", "请输入 ≥ 0 的数字")
            return
        }

        // 3. 应用
        let (updated, skipped) = setSelectionRadius(cm)
        if updated == 0 {
            showAlert("没改到东西", "选区里没有可调整的圆角矩形（可能形状类型不支持）")
        } else {
            var msg = "已更新 \(updated) 个圆角矩形的 R 角为 \(String(format: "%.2f", cm)) 厘米"
            if skipped > 0 { msg += "\n跳过 \(skipped) 个非圆角矩形" }
            showAlert("完成 ✓", msg)
        }
    }

    @objc func lockCurrent() {
        let info = readSelectionInfo()
        if info.isEmpty {
            showAlert("未选中圆角矩形", "请先在 PowerPoint 中选中一个或多个圆角矩形。")
            return
        }
        var locks = loadLocks()
        for s in info {
            locks[s.shapeId] = LockEntry(radiusCm: s.radiusCm, locked: true)
        }
        saveLocks(locks)
        showAlert("已锁定", "已锁定 \(info.count) 个圆角矩形的 R 角绝对值。\n改变形状大小后，点「重新应用锁定」即可恢复。")
    }

    @objc func unlockCurrent() {
        let info = readSelectionInfo()
        if info.isEmpty {
            showAlert("未选中", "未选中任何形状。")
            return
        }
        var locks = loadLocks()
        var n = 0
        for s in info {
            if locks.removeValue(forKey: s.shapeId) != nil { n += 1 }
        }
        saveLocks(locks)
        if n == 0 {
            showAlert("无需解锁", "选区里的 \(info.count) 个圆角矩形都没有锁定。")
        } else {
            showAlert("已解锁", "已从锁定表中移除 \(n) 个圆角矩形。")
        }
    }

    @objc func reapplyLocks() {
        let locks = loadLocks()
        if locks.isEmpty {
            showAlert("没有锁定", "锁定表为空，没有任何形状被锁定。")
            return
        }
        // 把所有锁定传给 AppleScript
        let (applied, notFound) = reapplyLocksById(locks)
        var msg = "已重新应用 \(applied) 个锁定"
        if notFound > 0 { msg += "\n（\(notFound) 个锁定在当前文档里找不到，可能已删除）" }
        showAlert("完成 ✓", msg)
    }

    @objc func revealLocks() {
        NSWorkspace.shared.activateFileViewerSelecting([locksFile])
    }

    @objc func showAbout() {
        let alert = NSAlert()
        alert.messageText = "R 角调整 v1.0"
        alert.informativeText = """
        PowerPoint 圆角矩形 R 角精确控制工具。

        使用方法：
        1. 在 PowerPoint 中选中圆角矩形
        2. 点此菜单栏图标 → 「调整 R 角...」输入厘米值
        3. 锁定 / 解锁：把 R 角绝对值固化，改变形状大小后用「重新应用锁定」恢复

        锁定信息存储在：
        \(locksFile.path)

        GitHub: github.com/Jerrrry666/radius_in_ppt
        """
        alert.runModal()
    }

    @objc func quit() {
        NSApp.terminate(nil)
    }

    // MARK: - UI 工具

    private func promptInput(title: String, message: String, placeholder: String, defaultValue: String) -> String? {
        let alert = NSAlert()
        alert.messageText = title
        alert.informativeText = message
        let input = NSTextField(frame: NSRect(x: 0, y: 0, width: 200, height: 24))
        input.placeholderString = placeholder
        input.stringValue = defaultValue
        input.alignment = .right
        input.font = NSFont.monospacedDigitSystemFont(ofSize: 14, weight: .regular)
        alert.accessoryView = input
        alert.addButton(withTitle: "应用")
        alert.addButton(withTitle: "取消")
        // 让输入框自动获得焦点
        DispatchQueue.main.async {
            alert.window.initialFirstResponder = input
            input.selectText(nil)
        }
        let result = alert.runModal()
        if result == .alertFirstButtonReturn {
            return input.stringValue
        }
        return nil
    }

    private func showAlert(_ title: String, _ message: String) {
        let alert = NSAlert()
        alert.messageText = title
        alert.informativeText = message
        alert.runModal()
    }

    // MARK: - 锁定文件读写

    private func loadLocks() -> LockMap {
        guard let data = try? Data(contentsOf: locksFile) else { return [:] }
        return (try? JSONDecoder().decode(LockMap.self, from: data)) ?? [:]
    }

    private func saveLocks(_ locks: LockMap) {
        let enc = JSONEncoder()
        enc.outputFormatting = [.prettyPrinted, .sortedKeys]
        if let data = try? enc.encode(locks) {
            try? data.write(to: locksFile, options: .atomic)
        }
    }

    // MARK: - AppleScript 桥接

    private func readSelectionInfo() -> [ShapeInfo] {
        let result = runAppleScript(Script.readSelectionInfo) ?? ""
        if result.isEmpty { return [] }
        return result
            .components(separatedBy: ";;" )
            .filter { !$0.isEmpty }
            .compactMap { line -> ShapeInfo? in
                let parts = line.components(separatedBy: "|")
                guard parts.count == 3,
                      let cm = Double(parts[1]),
                      let ss = Double(parts[2]) else { return nil }
                return ShapeInfo(shapeId: parts[0], radiusCm: cm, shortSide: ss)
            }
    }

    private func setSelectionRadius(_ cm: Double) -> (updated: Int, skipped: Int) {
        let script = Script.setSelectionRadius.replacingOccurrences(
            of: "{{CM}}",
            with: String(format: "%.6f", cm)
        )
        let out = runAppleScript(script) ?? "0|0"
        let parts = out.components(separatedBy: "|")
        return (Int(parts[0]) ?? 0, Int(parts[1]) ?? 0)
    }

    private func reapplyLocksById(_ locks: LockMap) -> (applied: Int, notFound: Int) {
        // 把 locks 序列化成 AppleScript list of records
        // 格式: {{shapeId1, radiusCm1}, {shapeId2, radiusCm2}, ...}
        var items: [String] = []
        for (id, entry) in locks {
            items.append("{\"\(escapeAS(id))\", \(String(format: "%.6f", entry.radiusCm))}")
        }
        let list = items.joined(separator: ", ")
        let script = Script.reapplyLocks.replacingOccurrences(of: "{{LIST}}", with: list)
        let out = runAppleScript(script) ?? "0|0"
        let parts = out.components(separatedBy: "|")
        return (Int(parts[0]) ?? 0, Int(parts[1]) ?? 0)
    }

    /// AppleScript 字符串内的双引号需要转义
    private func escapeAS(_ s: String) -> String {
        s.replacingOccurrences(of: "\"", with: "\\\"")
    }

    private func runAppleScript(_ source: String) -> String? {
        let tmp = URL(fileURLWithPath: "/tmp/radius_in_ppt_\(getpid())_\(Int.random(in: 1000...9999)).applescript")
        do {
            try source.write(to: tmp, atomically: true, encoding: .utf8)
        } catch {
            os_log("写临时 AppleScript 失败: %{public}@", log: log, type: .error, "\(error)")
            return nil
        }
        defer { try? FileManager.default.removeItem(at: tmp) }

        let task = Process()
        task.launchPath = "/usr/bin/osascript"
        task.arguments = [tmp.path]
        let pipe = Pipe()
        task.standardOutput = pipe
        task.standardError = pipe
        do {
            try task.run()
            task.waitUntilExit()
            let data = pipe.fileHandleForReading.readDataToEndOfFile()
            let out = String(data: data, encoding: .utf8)?
                .trimmingCharacters(in: .whitespacesAndNewlines)
            if task.terminationStatus != 0 {
                os_log("osascript 失败: %{public}@", log: log, type: .error, out ?? "(no output)")
            }
            return out
        } catch {
            os_log("osascript 执行失败: %{public}@", log: log, type: .error, "\(error)")
            return nil
        }
    }
}

// MARK: - AppleScript 源码

enum Script {
    /// 读取选区里所有圆角矩形的信息
    /// 返回: "id|radiusCm|shortSide;;id|radiusCm|shortSide;;..."
    static let readSelectionInfo = """
    tell application "PowerPoint"
        try
            set selShapes to selection
            set output to ""
            repeat with aShape in selShapes
                try
                    if (auto shape type of aShape) is rounded rectangle then
                        set shapeId to (id of aShape) as string
                        set w to width of aShape
                        set h to height of aShape
                        set shortSide to (w min h)
                        set ratio to (adjustment 1 of aShape)
                        set radiusCm to (ratio * shortSide) / 360000
                        if output is not "" then set output to output & ";;"
                        set output to output & shapeId & "|" & (radiusCm as string) & "|" & (shortSide as string)
                    end if
                end try
            end repeat
            return output
        on error errMsg
            return ""
        end try
    end tell
    """

    /// 设置选区里所有圆角矩形的 R 角绝对值
    /// {{CM}} 替换为厘米值
    /// 返回: "updated|skipped"
    static let setSelectionRadius = """
    on setRadius(targetCm)
        set updated to 0
        set skipped to 0
        tell application "PowerPoint"
            try
                set selShapes to selection
                repeat with aShape in selShapes
                    try
                        if (auto shape type of aShape) is rounded rectangle then
                            set w to width of aShape
                            set h to height of aShape
                            set shortSide to (w min h)
                            if shortSide > 0 then
                                set ratio to (targetCm * 360000) / shortSide
                                if ratio < 0 then set ratio to 0
                                if ratio > 0.5 then set ratio to 0.5
                                set adjustment 1 of aShape to ratio
                                set updated to updated + 1
                            else
                                set skipped to skipped + 1
                            end if
                        else
                            set skipped to skipped + 1
                        end if
                    on error
                        set skipped to skipped + 1
                    end try
                end repeat
            on error
                return "0|0"
            end try
        end tell
        return (updated as string) & "|" & (skipped as string)
    end setRadius
    setRadius({{CM}})
    """

    /// 按 id 在所有 slide 中找圆角矩形，重写 R 角
    /// {{LIST}} 替换为 AppleScript list of {id, cm}
    /// 返回: "applied|notFound"
    static let reapplyLocks = """
    on reapplyLocks(locksList)
        set applied to 0
        set notFound to 0
        tell application "PowerPoint"
            try
                set thePres to active presentation
                set theSlides to slides of thePres
                repeat with aLock in locksList
                    set targetId to (item 1 of aLock) as string
                    set targetCm to (item 2 of aLock) as number
                    set foundInSlide to false
                    repeat with aSlide in theSlides
                        set theShapes to shapes of aSlide
                        repeat with aShape in theShapes
                            try
                                if (id of aShape as string) is targetId then
                                    if (auto shape type of aShape) is rounded rectangle then
                                        set w to width of aShape
                                        set h to height of aShape
                                        set shortSide to (w min h)
                                        if shortSide > 0 then
                                            set ratio to (targetCm * 360000) / shortSide
                                            if ratio < 0 then set ratio to 0
                                            if ratio > 0.5 then set ratio to 0.5
                                            set adjustment 1 of aShape to ratio
                                            set applied to applied + 1
                                        end if
                                    end if
                                    set foundInSlide to true
                                    exit repeat
                                end if
                            end try
                        end repeat
                        if foundInSlide then exit repeat
                    end repeat
                    if not foundInSlide then set notFound to notFound + 1
                end repeat
            on error
                return "0|0"
            end try
        end tell
        return (applied as string) & "|" & (notFound as string)
    end reapplyLocks
    reapplyLocks({{LIST}})
    """
}

// MARK: - main

let app = NSApplication.shared
app.setActivationPolicy(.accessory)  // 不在 Dock 显示
let delegate = AppDelegate()
app.delegate = delegate
app.run()

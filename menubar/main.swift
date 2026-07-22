//
//  main.swift
//  R 角调整 — macOS 菜单栏 app（.pptx XML 路线）
//
//  工作流程（不用 AppleScript 读 selection，改用直接解析 .pptx）：
//    1. 用户点菜单栏「调整 R 角...」
//    2. Swift 用 AppleScript 拿当前 .pptx 文件路径
//    3. Swift 直接 unzip + 解析 slide*.xml，列出所有圆角矩形
//    4. 弹 NSWindow + NSTableView（多选 + 输入 R 角）
//    5. 用户选形状 + 输入 R 角 → 点「应用」
//    6. Swift 让 PowerPoint 保存 + 关闭
//    7. Swift 修改 .pptx XML（替换 <a:gd name="adj">）
//    8. Swift 让 PowerPoint 重新打开
//    9. 用户看到修改结果
//

import Cocoa
import os.log

// MARK: - 数据结构

struct ShapeEntry: Equatable {
    var id: String           // OOXML 形状 id
    var name: String         // 形状名
    var slide: Int           // slide 编号
    var shortSideCm: Double  // 短边（厘米）
    var currentRadiusCm: Double  // 当前 R 角（厘米）
    var ratio: Double        // 当前比例（0~0.5）
    var filePath: String     // 在 .pptx 内的 XML 路径，如 ppt/slides/slide1.xml
}

struct PptxDocument {
    let originalPath: String
    let workDir: String      // 解压后的临时目录
    let shapes: [ShapeEntry] // 所有圆角矩形
    let slideXmlPaths: [Int: String]  // slide# -> XML 路径
}

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
        statusItem = NSStatusBar.system.statusItem(withLength: NSStatusItem.variableLength)
        if let btn = statusItem.button {
            if let img = loadMenuBarIcon() {
                btn.image = img
            } else {
                btn.title = "R⤢"
            }
            btn.toolTip = "R 角调整 — 点击打开菜单"
        }
        statusItem.menu = buildMenu()
    }

    private func loadMenuBarIcon() -> NSImage? {
        if let path = Bundle.main.path(forResource: "menubar-icon", ofType: "png"),
           let img = NSImage(contentsOfFile: path) {
            img.size = NSSize(width: 18, height: 18)
            img.isTemplate = true
            return img
        }
        return nil
    }

    private func buildMenu() -> NSMenu {
        let menu = NSMenu()
        menu.autoenablesItems = false
        addMenu(menu, "调整 R 角...", action: #selector(adjustRadius), key: "r")
        menu.addItem(NSMenuItem.separator())
        addMenu(menu, "查看锁定文件", action: #selector(revealLocks), key: "")
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
        // 1. 拿到当前 .pptx 路径
        guard let pptxPath = getActivePptxPath() else {
            showAlert("找不到 PowerPoint 文档", "请先在 PowerPoint 中打开一个 .pptx 文件并保存。\n\n（注意：当前 .pptx 文件必须已经保存到磁盘，因为我们要直接修改文件）")
            return
        }
        guard FileManager.default.fileExists(atPath: pptxPath) else {
            showAlert("文件不存在", "PowerPoint 报告的路径：\n\(pptxPath)\n\n文件在磁盘上不存在。")
            return
        }

        // 2. 解析 .pptx，列出所有圆角矩形
        guard let doc = parsePptx(path: pptxPath) else {
            showAlert("解析失败", "无法解压 / 解析 .pptx 文件。")
            return
        }
        if doc.shapes.isEmpty {
            showAlert("没有圆角矩形", "这个 .pptx 文档里没有任何圆角矩形。")
            return
        }

        // 3. 弹选择窗口
        let dialog = ShapeSelectorController(document: doc) { [weak self] selectedShapes, cm in
            self?.applyRadius(pptxPath: pptxPath, selectedShapes: selectedShapes, cm: cm)
        }
        dialog.show()
    }

    @objc func revealLocks() {
        NSWorkspace.shared.activateFileViewerSelecting([locksFile])
    }

    @objc func showAbout() {
        let alert = NSAlert()
        alert.messageText = "R 角调整 v1.0"
        alert.informativeText = """
        PowerPoint 圆角矩形 R 角精确控制工具。

        工作原理（绕过 PowerPoint AppleScript bridge）：
        1. 拿到当前 .pptx 文件路径
        2. 解压 + 解析 slide*.xml
        3. 弹窗列出所有圆角矩形，你选要改的
        4. 保存关闭 PowerPoint
        5. 直接修改 .pptx 内的 XML
        6. 重新打开 PowerPoint

        锁定信息：\(locksFile.path)
        GitHub：github.com/Jerrrry666/radius_in_ppt
        """
        alert.runModal()
    }

    @objc func quit() {
        NSApp.terminate(nil)
    }

    // MARK: - 工具

    private func showAlert(_ title: String, _ message: String) {
        let alert = NSAlert()
        alert.messageText = title
        alert.informativeText = message
        alert.runModal()
    }

    // MARK: - 拿到当前 .pptx 路径

    private func getActivePptxPath() -> String? {
        let script = """
        tell application "PowerPoint"
            try
                if not (exists active window) then return ""
                return full name of (presentation of active window)
            on error
                return ""
            end try
        end tell
        """
        guard let out = runAppleScript(script), !out.isEmpty else { return nil }
        return out
    }

    private func runAppleScript(_ source: String) -> String? {
        let tmp = URL(fileURLWithPath: "/tmp/radius_in_ppt_\(getpid())_\(Int.random(in: 1000...9999)).applescript")
        do {
            try source.write(to: tmp, atomically: true, encoding: .utf8)
        } catch {
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
            return String(data: data, encoding: .utf8)?
                .trimmingCharacters(in: .whitespacesAndNewlines)
        } catch {
            return nil
        }
    }

    // MARK: - .pptx 解析

    private func parsePptx(path: String) -> PptxDocument? {
        let fm = FileManager.default
        let workDir = "/tmp/radius_in_ppt_work_\(getpid())_\(Int.random(in: 1000...9999))"
        try? fm.removeItem(atPath: workDir)
        try? fm.createDirectory(atPath: workDir, withIntermediateDirectories: true)

        // 用 ditto 或 unzip 解压
        let unzip = Process()
        unzip.launchPath = "/usr/bin/unzip"
        unzip.arguments = ["-o", "-q", path, "-d", workDir]
        let pipe = Pipe()
        unzip.standardOutput = pipe
        unzip.standardError = pipe
        do {
            try unzip.run()
            unzip.waitUntilExit()
            if unzip.terminationStatus != 0 {
                os_log("unzip 失败: status=%d", log: log, type: .error, unzip.terminationStatus)
                return nil
            }
        } catch {
            os_log("unzip 执行失败: %{public}@", log: log, type: .error, "\(error)")
            return nil
        }

        // 找所有 slide*.xml
        let slidesDir = (workDir as NSString).appendingPathComponent("ppt/slides")
        var slideXmlPaths: [Int: String] = [:]
        if let files = try? fm.contentsOfDirectory(atPath: slidesDir) {
            for f in files where f.hasPrefix("slide") && f.hasSuffix(".xml") {
                let numStr = f.dropFirst("slide".count).dropLast(".xml".count)
                if let num = Int(numStr) {
                    slideXmlPaths[num] = (slidesDir as NSString).appendingPathComponent(f)
                }
            }
        }

        // 解析每个 slide
        var shapes: [ShapeEntry] = []
        for (slideNum, xmlPath) in slideXmlPaths {
            guard let content = try? String(contentsOfFile: xmlPath, encoding: .utf8) else { continue }
            shapes.append(contentsOf: parseRoundedRects(xml: content, slide: slideNum, xmlPath: xmlPath))
        }

        return PptxDocument(
            originalPath: path,
            workDir: workDir,
            shapes: shapes,
            slideXmlPaths: slideXmlPaths
        )
    }

    /// 从单个 slide XML 解析所有圆角矩形
    private func parseRoundedRects(xml: String, slide: Int, xmlPath: String) -> [ShapeEntry] {
        var results: [ShapeEntry] = []
        // 找所有 <p:sp>...</p:sp>
        // 简单方法：找 "<p:sp>" 的位置，找匹配的 "</p:sp>"
        var idx = xml.startIndex
        while let spStartRange = xml.range(of: "<p:sp>", range: idx..<xml.endIndex) {
            guard let spEndRange = xml.range(of: "</p:sp>", range: spStartRange.upperBound..<xml.endIndex) else { break }
            let spXml = String(xml[spStartRange.lowerBound..<spEndRange.upperBound])
            if let entry = parseRoundedRect(spXml: spXml, slide: slide, xmlPath: xmlPath) {
                results.append(entry)
            }
            idx = spEndRange.upperBound
        }
        return results
    }

    private func parseRoundedRect(spXml: String, slide: Int, xmlPath: String) -> ShapeEntry? {
        // 检查 prstGeom prst="roundRect"
        guard spXml.contains("prst=\"roundRect\"") else { return nil }
        // 提取 id
        guard let idRange = spXml.range(of: "id=\"", range: spXml.startIndex..<spXml.endIndex),
              let idEnd = spXml.range(of: "\"", range: idRange.upperBound..<spXml.endIndex) else { return nil }
        let id = String(spXml[idRange.upperBound..<idEnd.lowerBound])
        // 提取 name（cNvPr 的 name）
        var name = ""
        if let nameAttr = spXml.range(of: "name=\"", range: spXml.startIndex..<spXml.endIndex) {
            if let nameEnd = spXml.range(of: "\"", range: nameAttr.upperBound..<spXml.endIndex) {
                name = String(spXml[nameAttr.upperBound..<nameEnd.lowerBound])
            }
        }
        // 提取 ext cx/cy
        var cx: Int = 0, cy: Int = 0
        if let extRange = spXml.range(of: "<a:ext ", range: spXml.startIndex..<spXml.endIndex) {
            if let extEnd = spXml.range(of: ">", range: extRange.upperBound..<spXml.endIndex) {
                let extXml = String(spXml[extRange.lowerBound..<extEnd.upperBound])
                cx = extractInt(from: extXml, attr: "cx") ?? 0
                cy = extractInt(from: extXml, attr: "cy") ?? 0
            }
        }
        // 提取 adj fmla 值
        var ratio: Double = 0
        if let adjRange = spXml.range(of: "name=\"adj\"", range: spXml.startIndex..<spXml.endIndex) {
            if let fmlaRange = spXml.range(of: "fmla=\"", range: adjRange.upperBound..<spXml.endIndex),
               let fmlaEnd = spXml.range(of: "\"", range: fmlaRange.upperBound..<spXml.endIndex) {
                let fmla = String(spXml[fmlaRange.upperBound..<fmlaEnd.lowerBound])
                // fmla 形如 "val 16952"，val 后面是 0~50000 之间的数（表示 0%~50%）
                if let valStr = fmla.components(separatedBy: " ").last,
                   let val = Double(valStr) {
                    ratio = val / 100000.0
                }
            }
        }
        let shortSideEmu = min(cx, cy)
        let shortSideCm = Double(shortSideEmu) * CM_PER_EMU
        let currentRadiusCm = shortSideEmu > 0 ? ratio * shortSideCm : 0
        return ShapeEntry(
            id: id, name: name, slide: slide,
            shortSideCm: shortSideCm, currentRadiusCm: currentRadiusCm,
            ratio: ratio, filePath: xmlPath
        )
    }

    private func extractInt(from xml: String, attr: String) -> Int? {
        let pattern = "\(attr)=\""
        if let r1 = xml.range(of: pattern),
           let r2 = xml.range(of: "\"", range: r1.upperBound..<xml.endIndex) {
            return Int(xml[r1.upperBound..<r2.lowerBound])
        }
        return nil
    }

    // MARK: - 应用 R 角

    private func applyRadius(pptxPath: String, selectedShapes: [ShapeEntry], cm: Double) {
        // 1. 让 PowerPoint 保存并关闭
        let saveScript = """
        tell application "PowerPoint"
            try
                save active presentation
                close active presentation saving no
                return "ok"
            on error errMsg
                return "err: " & errMsg
            end try
        end tell
        """
        let saveResult = runAppleScript(saveScript) ?? ""
        if !saveResult.contains("ok") {
            showAlert("保存失败", "无法保存并关闭 PowerPoint 文档。\n\n原因：\(saveResult)\n\n请确保 PowerPoint 当前文档是 .pptx 格式且已保存过。")
            return
        }

        // 2. 重新解压 + 修改
        guard let doc = parsePptx(path: pptxPath) else {
            showAlert("重新解压失败", "无法解压 .pptx 文件以修改。")
            return
        }

        // 按 (slide, id) 索引
        var byKey: [String: ShapeEntry] = [:]
        for s in doc.shapes { byKey["\(s.slide)|\(s.id)"] = s }

        var modified: [String] = []  // 修改过的 shape key
        for sel in selectedShapes {
            let key = "\(sel.slide)|\(sel.id)"
            guard let target = byKey[key] else { continue }
            let newRatio = cmToRatio(cm: cm, shortSideCm: target.shortSideCm)
            let newVal = Int(newRatio * 100000)
            // 修改 XML
            if modifyShapeAdj(xmlPath: target.filePath, shapeId: target.id, newVal: newVal) {
                modified.append(key)
            }
        }

        if modified.isEmpty {
            showAlert("没改到", "没有形状被修改。")
            return
        }

        // 3. 重新打包成 .pptx
        let repackResult = repackPptx(workDir: doc.workDir, outputPath: pptxPath)
        guard repackResult else {
            showAlert("打包失败", "无法把修改后的文件打包成 .pptx。")
            return
        }

        // 4. 重新打开 PowerPoint
        let openScript = """
        tell application "PowerPoint"
            try
                open POSIX file "\(pptxPath)"
                activate
                return "ok"
            on error errMsg
                return "err: " & errMsg
            end try
        end tell
        """
        let openResult = runAppleScript(openScript) ?? ""
        if !openResult.contains("ok") {
            showAlert("重开失败", "已修改 .pptx，但无法重新打开 PowerPoint。\n请手动打开：\n\(pptxPath)\n\n原因：\(openResult)")
            return
        }

        // 5. 写锁定（如果当前是锁定模式，暂不锁定，由用户后续点菜单锁定）
        showAlert("完成 ✓", "已更新 \(modified.count) 个圆角矩形的 R 角为 \(String(format: "%.2f", cm)) 厘米。\n\n提示：之后如需锁定 R 角绝对值，请用「查看锁定文件」里手动管理。")
    }

    /// 修改单个 shape 的 adj fmla 值
    private func modifyShapeAdj(xmlPath: String, shapeId: String, newVal: Int) -> Bool {
        guard var xml = try? String(contentsOfFile: xmlPath, encoding: .utf8) else { return false }
        // 找 id="X" 的那个 <p:sp>
        let idPattern = "id=\"\(shapeId)\""
        guard let idRange = xml.range(of: idPattern) else { return false }
        // 找这个 <p:sp> 的范围
        guard let spStart = xml.range(of: "<p:sp>", range: idRange.lowerBound..<xml.endIndex) else { return false }
        guard let spEnd = xml.range(of: "</p:sp>", range: spStart.upperBound..<xml.endIndex) else { return false }
        let spXml = String(xml[spStart.lowerBound..<spEnd.upperBound])
        // 确认是 roundRect
        guard spXml.contains("prst=\"roundRect\"") else { return false }
        // 替换 adj fmla
        let oldFmlaPattern = "name=\"adj\" fmla=\"val [0-9]+\""
        guard let oldRange = spXml.range(of: oldFmlaPattern, options: .regularExpression) else { return false }
        let newFmla = "name=\"adj\" fmla=\"val \(newVal)\""
        let newSpXml = spXml.replacingCharacters(in: oldRange, with: newFmla)
        xml = xml.replacingCharacters(in: spStart.lowerBound..<spEnd.upperBound, with: newSpXml)
        // 写回
        do {
            try xml.write(toFile: xmlPath, atomically: true, encoding: .utf8)
            return true
        } catch {
            return false
        }
    }

    /// 把 workDir 重新打包成 .pptx（覆盖 outputPath）
    private func repackPptx(workDir: String, outputPath: String) -> Bool {
        let fm = FileManager.default
        let tmpOut = "/tmp/radius_in_ppt_out_\(getpid())_\(Int.random(in: 1000...9999)).pptx"
        try? fm.removeItem(atPath: tmpOut)

        // 用 ditto 打包（保留 zip 结构）
        // ditto -c -k --sequesterRsrc --keepParent <dir> <archive>
        let ditto = Process()
        ditto.launchPath = "/usr/bin/ditto"
        ditto.arguments = ["-c", "-k", "--sequesterRsrc", "--keepParent", workDir, tmpOut]
        let pipe = Pipe()
        ditto.standardOutput = pipe
        ditto.standardError = pipe
        do {
            try ditto.run()
            ditto.waitUntilExit()
        } catch {
            return false
        }
        if ditto.terminationStatus != 0 {
            return false
        }
        // 覆盖
        do {
            try fm.removeItem(atPath: outputPath)
            try fm.copyItem(atPath: tmpOut, toPath: outputPath)
            try? fm.removeItem(atPath: tmpOut)
            return true
        } catch {
            return false
        }
    }
}

// MARK: - 形状选择窗口

class ShapeSelectorController: NSObject, NSTableViewDataSource, NSTableViewDelegate {
    let document: PptxDocument
    let onApply: ([ShapeEntry], Double) -> Void
    var window: NSWindow!
    var tableView: NSTableView!
    var checkboxColumn: NSTableColumn!
    var radiusField: NSTextField!
    var statusLabel: NSTextField!

    init(document: PptxDocument, onApply: @escaping ([ShapeEntry], Double) -> Void) {
        self.document = document
        self.onApply = onApply
        super.init()
    }

    func show() {
        // 主窗口
        let win = NSWindow(
            contentRect: NSRect(x: 0, y: 0, width: 600, height: 480),
            styleMask: [.titled, .closable, .miniaturizable],
            backing: .buffered, defer: false
        )
        win.title = "R 角调整 — \(URL(fileURLWithPath: document.originalPath).lastPathComponent)"
        win.center()

        let content = NSView(frame: win.contentView!.bounds)
        win.contentView = content
        self.window = win

        // 顶部说明
        let header = NSTextField(labelWithString: "勾选要调整的圆角矩形（按住 ⌘ 多选），然后输入 R 角值（厘米）：")
        header.frame = NSRect(x: 16, y: 440, width: 568, height: 24)
        header.font = NSFont.systemFont(ofSize: 12)
        content.addSubview(header)

        // 状态
        let status = NSTextField(labelWithString: "共 \(document.shapes.count) 个圆角矩形")
        status.frame = NSRect(x: 16, y: 416, width: 300, height: 18)
        status.font = NSFont.systemFont(ofSize: 11)
        status.textColor = .secondaryLabelColor
        content.addSubview(status)
        self.statusLabel = status

        // Table
        let scroll = NSScrollView(frame: NSRect(x: 16, y: 180, width: 568, height: 220))
        scroll.hasVerticalScroller = true
        scroll.borderType = .bezelBorder
        let table = NSTableView(frame: scroll.bounds)
        table.usesAlternatingRowBackgroundColors = true
        table.allowsMultipleSelection = true
        table.allowsEmptySelection = true
        table.delegate = self
        table.dataSource = self
        table.rowSizeStyle = .small
        // 列
        let col0 = NSTableColumn(identifier: NSUserInterfaceItemIdentifier("sel"))
        col0.title = ""
        col0.width = 30
        table.addTableColumn(col0)
        self.checkboxColumn = col0

        let col1 = NSTableColumn(identifier: NSUserInterfaceItemIdentifier("slide"))
        col1.title = "Slide"
        col1.width = 50
        table.addTableColumn(col1)

        let col2 = NSTableColumn(identifier: NSUserInterfaceItemIdentifier("name"))
        col2.title = "形状名"
        col2.width = 200
        table.addTableColumn(col2)

        let col3 = NSTableColumn(identifier: NSUserInterfaceItemIdentifier("current"))
        col3.title = "当前 R 角"
        col3.width = 100
        table.addTableColumn(col3)

        let col4 = NSTableColumn(identifier: NSUserInterfaceItemIdentifier("short"))
        col4.title = "短边 (cm)"
        col4.width = 100
        table.addTableColumn(col4)

        scroll.documentView = table
        content.addSubview(scroll)
        self.tableView = table

        // R 角输入
        let radiusLabel = NSTextField(labelWithString: "R 角（厘米）：")
        radiusLabel.frame = NSRect(x: 16, y: 130, width: 110, height: 22)
        radiusLabel.font = NSFont.systemFont(ofSize: 12)
        content.addSubview(radiusLabel)

        let field = NSTextField(frame: NSRect(x: 130, y: 128, width: 120, height: 24))
        field.stringValue = "0.30"
        field.alignment = .right
        field.font = NSFont.monospacedDigitSystemFont(ofSize: 14, weight: .regular)
        content.addSubview(field)
        self.radiusField = field

        let cmLabel = NSTextField(labelWithString: "厘米")
        cmLabel.frame = NSRect(x: 258, y: 130, width: 40, height: 22)
        cmLabel.font = NSFont.systemFont(ofSize: 12)
        cmLabel.textColor = .secondaryLabelColor
        content.addSubview(cmLabel)

        // 全选 / 全不选
        let selectAll = NSButton(title: "全选", target: self, action: #selector(selectAll))
        selectAll.frame = NSRect(x: 320, y: 128, width: 60, height: 24)
        selectAll.bezelStyle = .roundRect
        content.addSubview(selectAll)

        let deselectAll = NSButton(title: "全不选", target: self, action: #selector(deselectAll))
        deselectAll.frame = NSRect(x: 384, y: 128, width: 80, height: 24)
        deselectAll.bezelStyle = .roundRect
        content.addSubview(deselectAll)

        // 底部按钮
        let cancel = NSButton(title: "取消", target: self, action: #selector(cancel))
        cancel.frame = NSRect(x: 410, y: 30, width: 80, height: 32)
        cancel.bezelStyle = .roundRect
        content.addSubview(cancel)

        let apply = NSButton(title: "应用 R 角", target: self, action: #selector(apply))
        apply.frame = NSRect(x: 500, y: 30, width: 84, height: 32)
        apply.bezelStyle = .rounded
        apply.keyEquivalent = "\r"  // 回车
        content.addSubview(apply)

        // 提示
        let tip = NSTextField(labelWithString: "⚠️ 应用时会自动保存并关闭 PowerPoint 文档，修改 .pptx 文件后再重新打开。")
        tip.frame = NSRect(x: 16, y: 80, width: 568, height: 32)
        tip.font = NSFont.systemFont(ofSize: 10)
        tip.textColor = .systemOrange
        tip.maximumNumberOfLines = 2
        content.addSubview(tip)

        // 默认全选
        DispatchQueue.main.async { [weak self] in
            self?.selectAll()
        }

        NSApp.activate(ignoringOtherApps: true)
        win.makeKeyAndOrderFront(nil)
    }

    // MARK: - NSTableViewDataSource

    func numberOfRows(in tableView: NSTableView) -> Int {
        document.shapes.count
    }

    // MARK: - NSTableViewDelegate

    func tableView(_ tableView: NSTableView, viewFor tableColumn: NSTableColumn?, row: Int) -> NSView? {
        let shape = document.shapes[row]
        let identifier = tableColumn?.identifier.rawValue ?? ""
        let cell: NSTableCellView
        let cellIdentifier = NSUserInterfaceItemIdentifier("cell_\(identifier)")
        if let cached = tableView.makeView(withIdentifier: cellIdentifier, owner: self) as? NSTableCellView {
            cell = cached
        } else {
            cell = NSTableCellView()
            cell.identifier = cellIdentifier
            let tf = NSTextField(labelWithString: "")
            tf.translatesAutoresizingMaskIntoConstraints = true
            tf.frame = NSRect(x: 4, y: 0, width: (tableColumn?.width ?? 100) - 8, height: 18)
            tf.font = NSFont.systemFont(ofSize: 11)
            cell.addSubview(tf)
            cell.textField = tf
        }
        switch identifier {
        case "sel":
            cell.textField?.stringValue = tableView.isRowSelected(row) ? "☑" : "☐"
        case "slide":
            cell.textField?.stringValue = "\(shape.slide)"
        case "name":
            cell.textField?.stringValue = shape.name.isEmpty ? "(无名)" : shape.name
        case "current":
            cell.textField?.stringValue = String(format: "%.2f cm", shape.currentRadiusCm)
        case "short":
            cell.textField?.stringValue = String(format: "%.2f", shape.shortSideCm)
        default:
            cell.textField?.stringValue = ""
        }
        return cell
    }

    func tableView(_ tableView: NSTableView, didSelectRowIndexes indexes: IndexSet) {
        statusLabel.stringValue = "已选 \(indexes.count) / \(document.shapes.count)"
        // 刷新选择列
        for r in 0..<document.shapes.count {
            if let col = tableView.tableColumn(withIdentifier: NSUserInterfaceItemIdentifier("sel")) {
                tableView.reloadData(forRowIndexes: IndexSet(integer: r), columnIndexes: IndexSet(integer: col.index))
            }
        }
    }

    func tableView(_ tableView: NSTableView, didDeselectRowIndexes indexes: IndexSet) {
        statusLabel.stringValue = "已选 \(tableView.selectedRowIndexes.count) / \(document.shapes.count)"
        for r in 0..<document.shapes.count {
            if let col = tableView.tableColumn(withIdentifier: NSUserInterfaceItemIdentifier("sel")) {
                tableView.reloadData(forRowIndexes: IndexSet(integer: r), columnIndexes: IndexSet(integer: col.index))
            }
        }
    }

    // MARK: - 按钮

    @objc func selectAll() {
        let all = IndexSet(integersIn: 0..<document.shapes.count)
        tableView.selectRowIndexes(all, byExtendingSelection: false)
        statusLabel.stringValue = "已选 \(document.shapes.count) / \(document.shapes.count)"
    }

    @objc func deselectAll() {
        tableView.deselectAll(nil)
        statusLabel.stringValue = "已选 0 / \(document.shapes.count)"
    }

    @objc func cancel() {
        window.close()
    }

    @objc func apply() {
        let selectedRows = tableView.selectedRowIndexes
        if selectedRows.isEmpty {
            NSSound.beep()
            return
        }
        guard let cm = Double(radiusField.stringValue), cm >= 0 else {
            NSSound.beep()
            return
        }
        let shapes = selectedRows.map { document.shapes[$0] }
        window.close()
        onApply(shapes, cm)
    }
}

// MARK: - 工具常量

let CM_PER_EMU = 1.0 / 360000.0
let EMU_PER_CM = 360000.0

func cmToRatio(cm: Double, shortSideCm: Double) -> Double {
    if shortSideCm <= 0 || !cm.isFinite { return 0 }
    let ratio = (cm * EMU_PER_CM) / (shortSideCm * EMU_PER_CM)
    return max(0, min(0.5, ratio))
}

// MARK: - main

let app = NSApplication.shared
app.setActivationPolicy(.accessory)
let delegate = AppDelegate()
app.delegate = delegate
app.run()

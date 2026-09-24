// Invigoration's Battle.net sign-in window for macOS.
//
// Opens Battle.net's web sign-in in a WKWebView and watches every navigation decision itself.
// Battle.net finishes a successful sign-in by redirecting to http://localhost:0/?ST=<credential>,
// which arrives as a server-side redirect of the login form's submit. Avalonia's web view never
// surfaces that redirect, so its window spins forever; here it's caught however it arrives.
//
// Usage: invigoration-signin <url> [--title <text>] [--allow-any-host]
// Prints "ST=<credential>" on stdout and exits 0 once signed in. Exits 2 if the window is closed,
// 3 on a bad argument. Diagnostics (scheme, host and path only, never a query) go to stderr.
// --allow-any-host exists for Invigoration's own tests against a local page.

import AppKit
@preconcurrency import WebKit

let exitClosed: Int32 = 2
let exitBadArguments: Int32 = 3

func log(_ message: String) {
    FileHandle.standardError.write(("\(message)\n").data(using: .utf8)!)
}

func describe(_ url: URL?) -> String {
    guard let url else { return "(no address)" }
    let port = url.port.map { ":\($0)" } ?? ""
    let query = (url.query?.isEmpty == false) ? " (with query)" : ""
    return "\(url.scheme ?? "?")://\(url.host ?? "")\(port)\(url.path)\(query)"
}

/// The ST credential if this is Battle.net's final redirect to localhost.
func credential(in url: URL?) -> String? {
    guard let url, url.host == "localhost",
          let items = URLComponents(url: url, resolvingAgainstBaseURL: false)?.queryItems,
          let st = items.first(where: { $0.name == "ST" })?.value, !st.isEmpty
    else { return nil }
    return st
}

final class SignIn: NSObject, NSApplicationDelegate, NSWindowDelegate, WKNavigationDelegate, WKUIDelegate {
    let start: URL
    let title: String
    var window: NSWindow!
    var webView: WKWebView!
    var finished = false

    init(start: URL, title: String) {
        self.start = start
        self.title = title
    }

    /// Without a main menu, macOS has nowhere to route ⌘C/⌘V/⌘A, so pasting a password silently fails.
    func installMenus() {
        let main = NSMenu()

        let appItem = NSMenuItem()
        let appMenu = NSMenu()
        appMenu.addItem(withTitle: "Close Sign-In", action: #selector(NSApplication.terminate(_:)), keyEquivalent: "q")
        appItem.submenu = appMenu
        main.addItem(appItem)

        let editItem = NSMenuItem()
        let edit = NSMenu(title: "Edit")
        edit.addItem(withTitle: "Undo", action: Selector(("undo:")), keyEquivalent: "z")
        edit.addItem(withTitle: "Redo", action: Selector(("redo:")), keyEquivalent: "Z")
        edit.addItem(.separator())
        edit.addItem(withTitle: "Cut", action: #selector(NSText.cut(_:)), keyEquivalent: "x")
        edit.addItem(withTitle: "Copy", action: #selector(NSText.copy(_:)), keyEquivalent: "c")
        edit.addItem(withTitle: "Paste", action: #selector(NSText.paste(_:)), keyEquivalent: "v")
        edit.addItem(withTitle: "Select All", action: #selector(NSText.selectAll(_:)), keyEquivalent: "a")
        editItem.submenu = edit
        main.addItem(editItem)

        NSApp.mainMenu = main
    }

    func applicationDidFinishLaunching(_ notification: Notification) {
        installMenus()
        let configuration = WKWebViewConfiguration()
        // Battle.net's own cookies persist between sign-ins, so "Keep me logged in" can work.
        configuration.websiteDataStore = .default()

        webView = WKWebView(frame: NSRect(x: 0, y: 0, width: 600, height: 700), configuration: configuration)
        webView.navigationDelegate = self
        webView.uiDelegate = self

        window = NSWindow(
            contentRect: NSRect(x: 0, y: 0, width: 600, height: 700),
            styleMask: [.titled, .closable, .resizable, .miniaturizable],
            backing: .buffered,
            defer: false)
        window.title = title
        window.contentView = webView
        window.delegate = self
        window.center()
        window.makeKeyAndOrderFront(nil)
        NSApp.activate(ignoringOtherApps: true)

        log("opening \(describe(start))")
        webView.load(URLRequest(url: start))
    }

    /// Hands the credential back and quits. Only ever once.
    func finish(with st: String) {
        guard !finished else { return }
        finished = true
        log("caught the sign-in redirect")
        FileHandle.standardOutput.write("ST=\(st)\n".data(using: .utf8)!)
        exit(0)
    }

    // Every navigation, including the redirects of a form submit, is decided here.
    func webView(_ webView: WKWebView, decidePolicyFor navigationAction: WKNavigationAction,
                 decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        let url = navigationAction.request.url
        log("navigating to \(describe(url))")
        if let st = credential(in: url) {
            decisionHandler(.cancel)
            finish(with: st)
            return
        }
        decisionHandler(.allow)
    }

    func webView(_ webView: WKWebView, didReceiveServerRedirectForProvisionalNavigation navigation: WKNavigation!) {
        log("redirected to \(describe(webView.url))")
        if let st = credential(in: webView.url) {
            finish(with: st)
        }
    }

    func webView(_ webView: WKWebView, decidePolicyFor navigationResponse: WKNavigationResponse,
                 decisionHandler: @escaping (WKNavigationResponsePolicy) -> Void) {
        if let st = credential(in: navigationResponse.response.url) {
            decisionHandler(.cancel)
            finish(with: st)
            return
        }
        decisionHandler(.allow)
    }

    // If WebKit refuses localhost:0 before any of the above, the failing address still has it.
    func webView(_ webView: WKWebView, didFailProvisionalNavigation navigation: WKNavigation!, withError error: Error) {
        let failing = (error as NSError).userInfo[NSURLErrorFailingURLErrorKey] as? URL
        log("failed to load \(describe(failing)): \((error as NSError).code)")
        if let st = credential(in: failing) {
            finish(with: st)
        }
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        log("loaded \(describe(webView.url))")
    }

    // A page asking for a new window (a popup) is opened in this one instead.
    func webView(_ webView: WKWebView, createWebViewWith configuration: WKWebViewConfiguration,
                 for navigationAction: WKNavigationAction, windowFeatures: WKWindowFeatures) -> WKWebView? {
        log("popup requested for \(describe(navigationAction.request.url)); opening it here")
        if let st = credential(in: navigationAction.request.url) {
            finish(with: st)
        } else if navigationAction.request.url != nil {
            webView.load(navigationAction.request)
        }
        return nil
    }

    func applicationWillTerminate(_ notification: Notification) {
        if !finished {
            exit(exitClosed)
        }
    }

    func windowWillClose(_ notification: Notification) {
        log("window closed")
        exit(exitClosed)
    }
}

// Arguments
var arguments = Array(CommandLine.arguments.dropFirst())
var title = "Battle.net Sign-In"
var allowAnyHost = false
var address: String?
while !arguments.isEmpty {
    let next = arguments.removeFirst()
    switch next {
    case "--title" where !arguments.isEmpty: title = arguments.removeFirst()
    case "--allow-any-host": allowAnyHost = true
    default: address = next
    }
}

guard let address, let start = URL(string: address) else {
    log("usage: invigoration-signin <url> [--title <text>]")
    exit(exitBadArguments)
}

// Only ever Battle.net's own sign-in, over https, unless a test says otherwise.
if !allowAnyHost {
    guard start.scheme == "https", let host = start.host, host.hasSuffix(".account.battle.net") else {
        log("refusing to open \(describe(start)): not a Battle.net sign-in address")
        exit(exitBadArguments)
    }
}

// Quit if Invigoration goes away: it holds our stdin open, so end-of-file means it's gone.
FileHandle.standardInput.readabilityHandler = { handle in
    if handle.availableData.isEmpty {
        exit(exitClosed)
    }
}

let app = NSApplication.shared
let delegate = SignIn(start: start, title: title)
app.delegate = delegate
app.setActivationPolicy(.regular)
app.run()

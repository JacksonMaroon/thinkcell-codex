# Think-cell for Codex 0.3.0

By Jackson Maroon. A focused Codex plugin for donor-based native chart creation and scoped data updates.

## Install from the private repository

Clone [JacksonMaroon/thinkcell-codex](https://github.com/JacksonMaroon/thinkcell-codex). From its root folder:

```powershell
gh repo clone JacksonMaroon/thinkcell-codex
Set-Location thinkcell-codex
codex plugin marketplace add .
codex plugin add thinkcell-codex@thinkcell-codex-release
```

Start a new Codex task and ask it to use the thinkcell-edit skill. It will check the local runtime before an update. Requires Windows desktop PowerPoint, licensed think-cell and Python dependencies. Experimental naming is enabled by default on working copies. This repository is private and does not install the plugin globally.

For ordinary presentation creation, slide writing, and non-think-cell PowerPoint edits, install the companion [PowerPoint for Codex](https://github.com/JacksonMaroon/powerpoint-codex). Keep chart operations in this plugin; it does not add general slide-design or text/table-edit tools.

Read the [plugin guide](plugins/thinkcell-codex/README.md), [validation results](plugins/thinkcell-codex/VALIDATION.md), and [MIT license](plugins/thinkcell-codex/LICENSE). This is an experimental private distribution, not a public plugin-directory listing.

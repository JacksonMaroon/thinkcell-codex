# Think-cell for Codex

**Create chart slides from native think-cell donors and update their data with Codex.**

Created and maintained by **Jackson Maroon**, with Codex assistance. Version **0.4.0**, experimental desktop release.

This plugin contains one focused skill, `thinkcell-edit`, and local Python/PowerShell helpers. It uses think-cell's official JSON automation for data changes. An experimental namer identifies unnamed charts and assigns their automation names on working copies by default. A `--named-only` option disables that step.

## What it does

- Inspects native chart identity and embedded data without opening PowerPoint.
- Builds a data-request template from the chart, including its actual row layout.
- Creates native chart slides from user-supplied donors, optionally binding a company style file for new-element defaults.
- Adds an experimental finished-slide route: choose a chart and empty region, position the native donor through exclusive anchors, then preserve ordinary target content in a new slide copy.
- Updates bounded pie, bar/column/line/area/combination, scatter and bubble structures through the installed generator.
- Adds guarded native controls for secondary axes, series colors, scalar and percentage labels, CAGR endpoints, existing axis breaks, series connectors, selected Gantt edits, and existing Excel-link rebinding.
- Adds experimental fixed-structure waterfall and Mekko contracts, including calculated totals and independent widths.
- Checks intended values, embedded data, untouched sibling charts, theme and notes. Ordinary charts also use native-cache parity; waterfall/Mekko instead require specialized visual semantic review after exact model/datasheet checks.
- Saves/reopens an output in PowerPoint and renders a preview for visual review.
- Includes an optional whole-slide handoff for a separately connected PowerPoint plugin.

Creation uses a real donor slide. The [finished-slide route](skills/thinkcell-edit/references/existing-slide.md) supports bounded chartless targets and one-chart donors with native placement and preservation checks. The package does not generate arbitrary chart types from scratch, provide general slide design, or certify untested native features. Existing chart styles and features come from the donor. It rejects persistent or unknown external Excel links. The namer is a compatibility adapter, not an official think-cell naming API.

## Install in Codex

The private [GitHub repository](https://github.com/JacksonMaroon/thinkcell-codex) includes a local marketplace:

1. Clone the repository.
2. From its root folder, run:

   ```powershell
   gh repo clone JacksonMaroon/thinkcell-codex
   Set-Location thinkcell-codex
   codex plugin marketplace add .
   codex plugin add thinkcell-codex@thinkcell-codex-release
   ```

3. Start a new Codex task, then ask:

   > Use the thinkcell-edit skill to check my setup and update the existing think-cell chart with my new data.

You can also install from the configured marketplace using Codex's plugin browser. Installation copies the plugin into Codex's cache. The scripts still require the local runtime below. The repository does not install the plugin globally.

For ordinary presentation creation, slide writing, and non-think-cell PowerPoint edits, install the companion [PowerPoint for Codex](https://github.com/JacksonMaroon/powerpoint-codex). Keep chart data updates and native chart operations in this plugin. The companion does not substitute for this plugin's supported chart workflow.

If you want only the standalone skill, copy `skills/thinkcell-edit` from the plugin folder to a supported Codex skills directory. All runtime helpers are inside that skill. Keep only one installed copy enabled to avoid duplicate routing.

## Runtime requirements

- Windows desktop with Microsoft PowerPoint and licensed, connected think-cell.
- Python 3.10+ and the dependencies in `skills/thinkcell-edit/scripts/requirements.txt`.
- Windows PowerShell 5.1 for native Office and structured-storage operations.
- Optional: a connected PowerPoint plugin with native slide export and OOXML return tools.

The setup guide is [here](skills/thinkcell-edit/references/setup.md). There are no embedded API keys, bundled Office binaries, proprietary templates, client decks, or background services. The scripts make no network requests while processing charts. Data supplied to Codex or an Office connector remains subject to that product's settings and policies.

## Quick example

From the plugin folder, after installing the requirements into your chosen interpreter:

```powershell
python skills/thinkcell-edit/scripts/thinkcell.py doctor
python skills/thinkcell-edit/scripts/thinkcell.py inspect --input "D:/work/slide.pptx"
```

Use the returned slide/shape selectors to create a request. Codex can then edit the request from your supplied data and execute an update. See [usage](skills/thinkcell-edit/references/usage.md), [data format](skills/thinkcell-edit/references/data-contract.md), and [creation and features](skills/thinkcell-edit/references/creation-and-features.md). The command defaults to read-only preflight; `--execute` creates a separate output. Experimental naming remains enabled in both paths unless `--named-only` is supplied.

## Verification and limits

Read [VALIDATION.md](VALIDATION.md) for the release's actual checks. Passing automated checks is followed by visual review of the generated preview. It does not establish support for all chart layouts or prove atomic coauthoring safety.

The optional handoff needs an external PowerPoint connector. This package does not grant that access. If the connector is absent, the local file workflow still works. Chart generation and native verification may use a shared PowerPoint process; they do not quit the application or change unrelated presentations intentionally. Keep reports after a failure and inspect state before retrying.

## Efficient instructions

The skill loads only its entrypoint initially, then the needed data or handoff reference. Deterministic scripts handle parsing, naming and validation. Routine updates do not rerun the entire historical test suite. The package does not force a model, reasoning effort, or extra agents. Its instructions follow the current [OpenAI Astra guidance](https://developers.openai.com/api/docs/guides/latest-model) on clear task scope, instruction priority, concise communication and proportionate verification.

## Credit and distribution

Please credit **Jackson Maroon, Think-cell for Codex** when sharing or adapting the toolkit. Attribution metadata is provided in [CITATION.cff](CITATION.cff); the code is under the [MIT license](LICENSE). Dependencies retain their own licenses. Think-cell, Microsoft PowerPoint and OpenAI Codex are third-party products; this is an independent project and does not imply their endorsement.

The ZIP is a private release artifact, not proof of acceptance into a public plugin directory. Public listing requires a separate submission.

Official references: [think-cell JSON automation](https://www.think-cell.com/en/resources/manual/jsondataautomation), [element naming](https://www.think-cell.com/en/resources/manual/introductionautomation), [OpenAI plugin packaging](https://developers.openai.com/plugins/build/plugins), [skill authoring](https://learn.chatgpt.com/docs/build-skills).

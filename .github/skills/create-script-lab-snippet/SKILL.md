---
name: create-script-lab-snippet
description: 'Create or expand an Office.js Script Lab snippet from an API name. Use when adding sample coverage for an Excel, Word, PowerPoint, or Outlook API and its reference-documentation CSV mapping.'
---

# Create a Script Lab snippet

Create complete sample coverage for the Office.js API named by the user. The input is one API name, preferably in canonical form such as `Excel.Range.values`, `Word.Paragraph.insertText`, `PowerPoint.Shape`, or `Office.MessageCompose.subject`.

Do not assume that every API needs a new YAML file. First determine whether the clearest, smallest coverage is a new sample or an expansion of an existing sample.

## Required outcome

Complete all applicable parts of the change:

- Add or expand a public snippet under `samples/<host>/<group>/`.
- Give the sample deterministic data in a workbook, document, presentation, or Outlook item when the host permits it; otherwise state exact item prerequisites.
- Add rows to the host file in `snippet-extractor-metadata/` that map the demonstrated API members to named snippet functions.
- Make the Office.js object model and types clear inside every mapped function.
- Regenerate repository outputs with the existing build.
- Validate the snippet with the repository's existing tests.

Do not finish with only a proposed sample or a coverage recommendation.

## 1. Resolve the API

Normalize the input into:

- package or namespace, such as `Excel`, `Word`, `PowerPoint`, or `Office`;
- Office host, including Outlook when an API is exposed through the `Office` namespace;
- class, interface, enum, or type name;
- member name, when one was supplied;
- overload number or top-level category (`class`, `interface`, `enum`, or `type`);
- API requirement set and whether the API is stable or preview.

Use authoritative Office.js API documentation when repository evidence is insufficient. If the user gives an unqualified name, resolve it from the repository and documentation. Ask only when multiple valid APIs remain genuinely indistinguishable.

Use the stable Office.js library unless the requested API is preview-only. Set `api_set` to the highest requirement set used anywhere in the sample.

For Outlook, also resolve the required activation context:

- Message Read
- Message Compose
- Appointment Attendee
- Appointment Organizer
- a supported combination of those modes

Confirm any client, Exchange account, mailbox, permission, or feature-specific limitation documented for the API.

## 2. Audit existing coverage

Search both of these surfaces before editing:

1. The appropriate CSV:
   - `snippet-extractor-metadata/excel.csv`
   - `snippet-extractor-metadata/word.csv`
   - `snippet-extractor-metadata/powerpoint.csv`
   - `snippet-extractor-metadata/outlook.csv`
2. All YAML files under `samples/` for the API namespace, class, member, related object-model access path, and nearby feature terminology.

Treat these cases separately:

- A CSV row exists and its mapped function clearly demonstrates the API: improve it only if it fails the requirements in this skill.
- A relevant snippet uses the API but has no suitable CSV mapping: expand or clarify that snippet and add the mapping.
- A related snippet can accept one focused action without mixing unrelated scenarios: expand it.
- No coherent sample exists, or expansion would make an existing sample broad or confusing: create a new snippet.

Prefer expansion when the requested API uses the same setup data, object model, requirement set, and user scenario as an existing sample. Prefer a new snippet when it needs substantially different setup, belongs to another gallery group, changes stable versus preview placement, or would make the existing sample difficult to understand.

Do not duplicate an existing scenario merely to create a one-API file.

## 3. Design the sample

Each sample must be understandable and repeatable. Document-host samples must not rely on the user's current file. Outlook samples must clearly define the mailbox item and activation context they require.

### Setup requirements

For Excel, Word, and PowerPoint, add a visible **Set up** action that creates deterministic data and named artifacts for subsequent actions.

- **Excel:** Create or replace a dedicated worksheet, normally named `Sample`. Add the ranges, tables, charts, or other objects needed by the API, format them when useful, activate the sheet, and call `context.sync()`.
- **Word:** Create a known sample document state with the paragraphs, ranges, tables, content controls, shapes, or other objects needed by the API. Do not depend on arbitrary existing content or selection unless selection is the API being demonstrated.
- **PowerPoint:** Create known slides and shapes, text, images, or other objects needed by the API. Do not depend on arbitrary existing slides or selection unless selection is the API being demonstrated.

Make setup safe to run repeatedly. Reuse or replace clearly named sample artifacts instead of accumulating ambiguous duplicate content.

Keep setup separate from the focused API actions unless the API itself creates the sample object. Do not map generic setup plumbing to an API merely to increase coverage.

### Outlook context and setup

Outlook does not provide a general equivalent of creating a sample workbook, document, or presentation. Follow the established Outlook pattern instead:

- State **Required mode** prominently in the HTML introduction. Include the item type when it matters, such as `Message Compose` or `Appointment Organizer`, rather than only `Compose` or `Read`.
- Use `api_set: Mailbox: '<version>'` for stable APIs and `Mailbox: preview` for preview APIs.
- Place preview-only samples under `samples/outlook/99-preview-apis/`.
- For Compose APIs, operate on the current draft item. Add a **Set up** action when the API needs deterministic subject, body, recipients, attendees, attachments, recurrence, or other mutable state. Make repeated setup safe.
- For Read APIs, do not pretend that the snippet can rewrite the current item. Provide concise numbered prerequisites that tell the user what message or appointment to create, send, receive, open, or select.
- When the API itself displays a new message or appointment form, use deterministic values in the form options so the displayed item is the sample data.
- For APIs driven by selection, item multi-select, drag-and-drop, or mailbox events, use `Office.onReady` when registration must happen at initialization. Explain the exact user action, supported mode, and any client limitation in the HTML.
- Do not combine Read and Compose behavior merely because both are Outlook. Reuse one sample only when the same API, code path, requirement set, and instructions are genuinely clear in every listed mode.
- Never send an item, close a compose form, or make another disruptive change as setup for an unrelated API.

### Function design

Use one clearly named function per focused operation. Split a workflow into as many meaningful, independently runnable actions as the API supports, such as prepare or mark entries, create, get properties, set properties, update, and delete. Give each action its own clearly labeled button and present the buttons in workflow order.

Always keep get and set operations in separate functions. A create or mutation function may log a short completion message, but move property loading and detailed console output into a dedicated get function. Do not combine steps merely to reduce the number of functions; discrete functions are easier to understand and can provide distinct reference-documentation examples.

A successful action must give the user an unambiguous way to verify its effect. When an API changes hidden document state, such as inserting field codes, explain that the visible document might not change and identify the next action or host setting that proves the operation succeeded. Keep this guidance in the task pane when it is necessary to understand the workflow.

Do not log an Office.js client object directly. An unloaded proxy commonly appears as an empty object in the console and does not help the user verify the action. Mutation functions should log a concise completion message. Dedicated get functions should explicitly load the relevant properties, synchronize, and then log those values.

A function may support multiple CSV rows only when it genuinely demonstrates closely related APIs on different reference pages.

Every function that will be mapped in a CSV must remain understandable when extracted from the full snippet:

- Put a short summary comment immediately inside the outer mapped function so the extracted sample includes it. Rephrase the operation that follows in plain English instead of repeating the code syntax.
- Add another short comment before a distinct step only when it helps the reader follow the sample. Keep comments concise, use active voice and present tense, and follow the [Microsoft Writing Style Guide](https://learn.microsoft.com/en-us/style-guide/).
- Show the access path from `context.workbook`, `context.document`, `context.presentation`, or `Office.context.mailbox` to the focal object.
- Add explicit Office.js type annotations to focal objects, collections, option objects, result objects, and event arguments.
- Use descriptive names that identify the object, such as `salesTable: Excel.Table`, `firstParagraph: Word.Paragraph`, `titleShape: PowerPoint.Shape`, or `message: Office.MessageCompose`.
- For an interface or type mapping, declare an object with that exact type.
- For an enum mapping, declare a value with that exact enum type when practical and use it in the API call.
- For a property or method mapping, type the owning object so the class-member relationship is obvious.
- Avoid `any`, unnecessary assertions, unexplained helper indirection, and generic names such as `item`, `object`, or `result` for the focal API.
- For Excel, Word, and PowerPoint, keep required `load` and `context.sync()` calls visible and in the correct order.
- Log a concise result or visibly update the host document so the action's effect can be verified.

Do not hide the focal API access inside an unmapped helper. Avoid dependencies on globals or sibling helpers in a mapped function unless the extracted function remains useful and the object model remains obvious without them.

For Outlook mapped functions:

- Assign `Office.context.mailbox.item` to a descriptively named variable with the exact applicable interface, such as `Office.MessageCompose`, `Office.MessageRead`, `Office.AppointmentCompose`, or `Office.AppointmentRead`. Narrow or assert the mailbox item union once when TypeScript requires it.
- Type focal child objects and options with their Office.js interfaces, such as `Office.Body`, `Office.Recipients`, `Office.Recurrence`, or the documented options interface.
- Type callback results as `Office.AsyncResult<T>` and event parameters with the documented event-argument interface when the definitions support it.
- Check `asyncResult.status` before reading `asyncResult.value`. Log `asyncResult.error.message` and return on failure.
- Keep the focal asynchronous call, callback result handling, and visible outcome together in the mapped function.
- Use `Office.MailboxEnums` values rather than equivalent string literals when the API accepts the enum.
- Do not use `Excel.run`, `Word.run`, `PowerPoint.run`, `load`, or `context.sync()` patterns for Outlook mailbox APIs.
- Do not use `as any` to bypass the Outlook item union or incomplete option typing. Use the exact Read/Compose and Message/Appointment interface instead.

## 4. Follow repository conventions

Before writing code, inspect the closest samples in the same host and group and follow their current patterns.

- Use a kebab-case `.yaml` filename in the most specific existing group.
- Create a new group only if no existing group fits. Preserve that host's numbering convention.
- Do not manually invent a new snippet `id`; let the build generate it.
- Use TypeScript, four-space indentation after YAML block indentation, double quotes in snippet code, and semicolons.
- For Excel, Word, and PowerPoint, register buttons with `document.getElementById(...).addEventListener("click", () => tryCatch(...))` and include the repository's standard `tryCatch` helper.
- For Outlook, follow neighboring samples by registering callback-based actions directly, such as `document.getElementById("get").addEventListener("click", get)`. Use `Office.onReady` for initialization and event registration when required.
- Use lowercase hyphenated HTML IDs. For document hosts, include a separate **Set up** section followed by **Try it out** actions. For Outlook, include **Set up** only when the sample can safely prepare its current item.
- Do not add visible step numbers to task-pane instructions or action labels when button order already makes the workflow clear. Use a numbered procedure only when the user must perform a strict sequence outside the task pane.
- Keep the name and description concise and specific to the demonstrated behavior.
- Use the canonical stable or beta Office.js URL and matching type definitions already used by neighboring snippets.
- For Outlook, use the canonical `https://officeapis.public.onecdn.static.microsoft/1/office.js` library used by neighboring samples.
- Do not introduce unrelated libraries or abstractions.

### Order snippets within the group

Read the `name` and `order` metadata from every snippet in the target group before choosing the new or expanded sample's position.

Use this display order:

1. Foundational samples whose names identify them as **Basic** samples come first.
2. Sort all remaining samples alphabetically by the YAML `name` value, using a case-insensitive comparison.

Do not classify a sample as basic merely to move it earlier. Its name and scenario must genuinely present the group's introductory or foundational API usage.

The playlist build sorts by `group`, then `order`, then `id`; it does not sort by `name`. If the group uses `order`, assign or update numeric `order` values so the complete group follows the required sequence. If the group has no `order` values and its ID ordering would not produce this sequence, add sequential `order` values to every snippet in that group. Avoid duplicate order values and preserve intentional relative ordering among multiple Basic samples unless their names make a clearer order necessary.

For an existing snippet, preserve its established formatting unless the build normalizes it.

## 5. Add documentation mappings

After a new snippet has been processed once by the build, read its generated `id`. Add mapping rows to the corresponding host CSV using this schema:

```csv
Package,Class,Member Name,Member ID or top-level category,SnippetIdInTheYAMLFile,MethodNameInTheSnippet
Excel,Range,values,,excel-range-set-get-values,setValues
Excel,Range,insert,1,excel-range-insert-delete-clear-range,insertShiftDown
Excel,CalculationMode,,enum,excel-workbook-calculation,switchToManualCalculations
Office,MessageCompose,subject,,outlook-other-item-apis-get-set-subject-compose,get
```

Apply these rules:

- Preserve the CSV header, UTF-8 BOM if present, line endings, quoting, and existing sort order.
- Use the exact package, class, and member spelling from the API documentation.
- For methods, specify the documented overload number in `Member ID or top-level category`.
- For properties, leave that column empty.
- For a top-level class, interface, enum, or type mapping, leave `Member Name` empty and put the category in `Member ID or top-level category`.
- Use the exact generated snippet `id`.
- Use the exact TypeScript function name; mapped functions must use a `function name(...)` declaration that `config/build.documentation.ts` can extract.
- Map a function at most once to any single API reference webpage. Do not map the same function to a class and several of that class's properties or methods, because those mappings repeat identical code on the same class page.
- If one function demonstrates several members on the same page, choose the single row that best represents its purpose. For a requested class, prefer the top-level `class` row. For a requested member, prefer that exact member row.
- The same function may appear in multiple rows only when each row targets a genuinely different API page, such as a class page, its collection page, and an options-interface page.
- Different functions may map to the same API page when they provide distinct, useful examples rather than duplicated code.
- Do not map incidental APIs used only for setup, navigation, loading, synchronization, logging, or cleanup.
- Do not add duplicate rows for coverage already provided by the same function.

If a mapped function is renamed or moved to another snippet, update every affected CSV row in the same change.

## 6. Add or update test coverage

Do not treat running the existing tests as sufficient. Determine how each new or expanded snippet is covered and update the test infrastructure when needed.

### Compilation coverage

`tests/snippet-compiler.test.ts` discovers snippet YAML files automatically. Every snippet must compile against the correct stable or preview Office.js definitions.

- Run `npm run test:compile` after the snippet code and libraries are final.
- Fix source or type errors in the snippet rather than weakening types.
- Change shared compiler declarations only when the authoritative Office.js definitions require support that the test harness genuinely lacks.

### Runtime coverage

Runtime tests are smoke tests for syntax, API names, registered handlers, and basic execution. They do not prove real Office behavior, collection mutation, `load`/`sync` semantics, ordering, visual output, or error conditions.

For every new sample or group:

1. Inspect `tests/runtime-auto-generated.test.ts`, especially `INCLUDED_GROUPS`, `EXCLUDED_PATTERNS`, and its button-clicking strategy.
2. Determine whether the snippet can execute meaningfully with the current host mock.
3. If a stable new group is mockable, add the group to `INCLUDED_GROUPS`.
4. If the snippet needs missing Office.js objects, methods, collections, enums, or result values, extend the appropriate factory in `tests/helpers/mock-factories.ts` with the smallest faithful mock.
5. If auto-generated coverage cannot exercise the focal action, add a focused case to `tests/runtime-execution.test.ts` using the existing snippet test runners.
6. If relying on auto-generated coverage, ensure the primary workflow uses the recognized `setup` and `run` button IDs. Do not rename a clearer public action solely for the harness; use a focused test instead when that produces better sample UX.
7. Add an exclusion only when the Office behavior cannot be represented meaningfully by the existing smoke-test architecture. Keep the exclusion narrow and document the concrete limitation next to it.

Do not add success-shaped mocks that conceal invalid code, and do not assert behavior the mock cannot model. Passing runtime tests never replaces manual testing in the real Office host.

Run the smallest applicable runtime command while iterating:

- `npm run test:runtime:auto` for an included group or exclusion change.
- `npm run test:runtime` for a focused runtime test.
- `npx jest <test-file-pattern>` when only one test file needs to run.

All snippets must still be manually tested in Script Lab before submission. State this requirement when handing off a snippet that has not been tested in the real host.

## 7. Build and validate

Use the repository's existing commands in this order:

1. For a new snippet, run `npm run build` once so the build writes its `id`.
2. Add or finish the CSV mappings.
3. Add or update compilation and runtime test coverage as described above.
4. Run `npm start` to compile configuration code, process snippets, regenerate playlists and views, generate documentation extracts, and lint.
5. Run `npm test`.

If a targeted test exposes a problem, fix it and rerun the smallest applicable test. Before finishing, run the complete required commands above successfully.

Review the generated documentation entry in `snippet-extractor-output/snippets.yaml`. Confirm that:

- every new CSV row produced an entry with the intended API key;
- the extracted code comes from the intended function;
- the extracted function makes the focal object type and class-member relationship obvious;
- the link points to the correct snippet;
- no generated mapping or playlist change is missing.

Retain all build-generated changes that correspond to the edited snippet, including the host playlist. Do not hand-edit generated playlist, view, or extractor output to conceal a source or mapping problem.

Finally, inspect the diff for unrelated generated churn and remove only changes caused by an accidental local operation. Never discard pre-existing user changes.

# Repository Guidance for Agents

## Communication

- Write all GitHub issue comments, pull request comments, pull request descriptions, commit messages, and other GitHub-facing text in English.
- Keep user-facing explanations in the language used by the user unless asked otherwise.

## Public and Protected API

- Preserve binary/source compatibility whenever possible. Do not remove or rename public members without an explicit request.
- When replacing an existing public member, keep the old member as an obsolete wrapper and delegate to the new implementation with the previous default behavior.
- Keep API wording consistent with existing naming. For example, prefer existing `Remove...` terminology over introducing `Delete...` for the same conceptual operation.
- Avoid unrelated API changes or side features while implementing a requested feature.

## XML Documentation

- Every new public API member must have XML documentation matching the quality of the surrounding API.
- New protected members that define engine contracts or are intended for derived engine implementations must also have XML documentation.
- Public enums and their values must be documented.
- Keep `<summary>` text short and focused, for example `Insert one or more columns.`.
- Put details such as zero-based indexing, insertion position, behavior contracts, and parameter semantics into `<param>`, `<returns>`, `<remarks>`, or `<exception>` elements as appropriate.
- Prefer simple grammar that is easy to understand for non-native English speakers.
- For overrides, use `<inheritdoc/>` when the base documentation is sufficient. Add `<remarks>` only for engine-specific behavior or limitations.

## Excel Engine Behavior

- Engine implementations must document and enforce whether formulas and references are automatically updated during structural workbook changes.
- If an engine can only support one formula/reference update mode, reject unsupported requests with a clear exception.
- Preserve engine-specific limitations explicitly and keep the original exception as an inner exception when wrapping known unsupported workbook structures.

## Tests and Generated Files

- Add durable unit tests for new behavior, including engine-specific behavior where engines differ.
- Test methods should include short comments or XML summaries explaining why the test exists and what workbook behavior it verifies.
- Static test workbooks belong in the appropriate `test_data` directories.
- When repository copy/clone scripts generate or synchronize shared source or test files, include the resulting copied files in the same change.

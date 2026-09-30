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
- The documentation requirements do not apply to the `EPPlus45-FixCalcsEdition.MultiTarget` project or to `TestAndDemoExcelOps`; do not make documentation-only changes there.
- Historical documentation baselines should stay at zero. If a member is intentionally out of scope, add a narrow checker exclusion instead of raising a global allowance.
- Public enums and their values must be documented.
- Keep `<summary>` text short and focused, for example `Insert one or more columns.`.
- For new or substantially edited XML documentation, prefer Microsoft-style complete sentences in descriptive third-person form with a final period, for example `Inserts one or more columns.`.
- Treat broad cleanup of older XML documentation style as a separate follow-up task instead of mixing it into feature or targeted documentation commits.
- Put details such as zero-based indexing, insertion position, behavior contracts, and parameter semantics into `<param>`, `<returns>`, `<remarks>`, or `<exception>` elements as appropriate.
- Prefer simple grammar that is easy to understand for non-native English speakers.
- For overrides, always add an explicit `<inheritdoc/>` when the inherited documentation applies. Do not rely on implicit inherited documentation from an undocumented override.
- For overloads with mostly identical documentation, prefer `<inheritdoc/>` plus targeted `<param>`, `<returns>`, `<remarks>`, or `<exception>` overrides instead of copying large documentation blocks.
- Add `<remarks>` to inherited documentation only for engine-specific behavior or limitations.
- Run `./tools/check-api-docs.ps1` before committing API changes. The script enforces the current documentation baseline and prevents new undocumented public/protected API members or additional overrides without `<inheritdoc/>`.

## Excel Engine Behavior

- Engine implementations must document and enforce whether formulas and references are automatically updated during structural workbook changes.
- If an engine can only support one formula/reference update mode, reject unsupported requests with a clear exception.
- Preserve engine-specific limitations explicitly and keep the original exception as an inner exception when wrapping known unsupported workbook structures.

## Tests and Generated Files

- Add durable unit tests for new behavior, including engine-specific behavior where engines differ.
- Test methods should preferably include short comments or XML summaries explaining why the test exists and what workbook behavior it verifies, but this is guidance and not a mandatory API documentation requirement.
- Static test workbooks belong in the appropriate `test_data` directories.
- When repository copy/clone scripts generate or synchronize shared source or test files, include the resulting copied files in the same change.

## Release Process

- Create releases only from the repository's primary integration branch, currently `main` or `master`.
- Do not create a release from a feature branch, ticket branch, Codex branch, or any other branch that has not been merged into the primary integration branch.
- Before creating a release, ensure the pull request has been created, reviewed as required, merged into the primary integration branch, and the build-and-test workflow for that branch has completed successfully.
- If the workflow is configured to run on pull requests and on pushes to the primary integration branch, wait for the relevant successful run after merge before creating the release.
- If a release is requested before these prerequisites are met, create or update the pull request first and explicitly tell the user that the release must wait for the successful build-and-test pipeline on the primary integration branch.
- Every release description must identify the affected project or projects for each listed change. Use the exact NuGet package ID where available; otherwise, use the project or repository-infrastructure name.
- Derive the affected-project list from the complete diff since the previous release. List all synchronized or generated project variants that changed, and label repository-wide tooling or documentation changes explicitly instead of leaving their scope ambiguous.

## Branch Cleanup

- After a pull request is merged, delete its feature, ticket, or Codex branch both locally and on the remote only after all associated pipelines have completed successfully, including pull-request checks, post-merge builds and tests, and any requested release or deployment workflows.
- Keep branch references available while any associated workflow is queued, running, waiting for approval, or requires a retry. Premature deletion can cause checkout failures.
- Before cleanup, verify that the exact branch tip is fully integrated into the primary integration branch and that no uncommitted work or unmerged commits would be lost.
- Switch the primary worktree back to the primary integration branch and update it by fast-forward. Safely detach any other clean worktree that still has the feature branch checked out before deleting the branch. Do not remove worktrees or their files without authorization.
- If a pipeline fails or its status is uncertain, retain the branch and report the outstanding condition instead of treating cleanup as complete.

## File Encoding and Line Endings

- Save text files as UTF-8 with BOM and CRLF line endings, matching `.editorconfig`.
- Keep `.gitattributes` line-ending rules intact and mark binary workbook/image/archive formats as binary.
- When normalizing encoding or line endings, keep that work in a separate mechanical commit whenever possible.

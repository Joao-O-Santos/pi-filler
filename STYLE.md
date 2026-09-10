# Style

These conventions keep the package small and easy to inspect.

## Code

Use tabs for indentation in JavaScript and TypeScript. Use double quotes
and let Biome organize imports. The formatter ceiling is 100 characters,
not a target: most lines should remain naturally shorter when that makes
the code clearer.

Prefer direct control flow and ordinary data structures. Keep functions
small enough to understand locally, but do not split code merely to make
functions shorter. A descriptive local variable is usually preferable
to a helper used once.

Use Node.js and platform APIs directly when they already provide the
needed behavior. Add dependencies when they remove meaningful work or
risk, not merely to avoid writing a few obvious lines.

## Markdown

Hard-wrap prose at 72 characters. Run `npm run markdown:fix` rather than
manually chasing wrapping differences. Leave URLs, tables, code blocks,
headings, and other syntax intact when wrapping them would reduce
readability or validity.

Keep documentation concise and current. Prefer a short durable document
to an accumulating history of implementation discussion.

## Naming

Use ordinary descriptive names. Prefer names that describe the thing's
role rather than its implementation mechanism. Keep public interfaces
small and unsurprising.

## Commits

Write commit subjects as imperative instructions describing the change:

```text
Add PDF search
Simplify process handling
Remove obsolete compatibility path
```

Keep the subject within 50 characters. Begin with a capitalized verb in
present imperative form. Do not write `Added`, `Adding`, or conventional
commit prefixes such as `fix:`, `feat:`, `chore:`, or `docs:`.

Use a body only when the reason or trade-off is not obvious from the
diff. Wrap commit-body prose at 72 characters too.

## Dependencies

Follow current upstream releases rather than freezing the project around
old versions. Do not commit a package lock. When current upstream breaks
something, adapt the project and its tests.

Markdown that is not a template, and is not declared as an AdditionalFile.

TargetsTests reads it back out of `@(None)`: the package drops the None identity the SDK's default
glob gives a template, and this file is what proves the drop is scoped to templates rather than to
every `.md` in the project.

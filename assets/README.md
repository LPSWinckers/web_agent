# Brand assets

The Icon Composer projects in `dev/`, `nightly/`, and `prod/` are the source for T3 Code's web
icons. Run `vp run icons:export` to regenerate them and copy the development favicon and splash
assets into `apps/web/public`. Use `vp run icons:check` to compare the generated files without
changing them.

The exporter still writes unused native-platform PNG and ICO files. The desktop and mobile apps no
longer consume them; pruning those outputs from the exporter and repository remains a follow-up.

Do not edit generated PNG or ICO files directly. Change the Icon Composer source and rerun the
exporter instead.

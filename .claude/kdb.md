# kdb+/q reference

Conventions and gotchas for working with the `.q` code in this repo.

## Critical language rules

- **Right-to-left evaluation, no operator precedence.** `2*3+4` is `2*(3+4)`
  = 14, not 10.
- **`%` is division, not modulo.** Use `mod` for modulo.
- **`:` is assignment, `=` is equality.** `x=42` compares; it doesn't assign.
- **`=` is element-wise, `~` is structural match.** Never use `=` to compare
  two lists for equality — `(1 2 3)=(1 2 4)` is a boolean vector, not a
  single bool. Use `~`.
- **No negative indexing.** `x -1` is subtraction. Use `last x`.
- **Atoms and vectors are different types.** `5` is a long atom (`-7h`),
  `enlist 5` is a one-element long vector (`7h`). `each` over a plain atom
  still applies the function once to that atom — this repo's `.xls.fmt`
  relies on exactly this to handle both a single table name and a list of
  table names with the same code path.
- **`if[]` has no return value.** Use `$[cond;true;false]` in expressions.
- **Reserved words** (`type`, `string`, `key`, `value`, `count`, `sum`,
  `where`, `in`, `get`, `set`, `abs`, `first`, `last`, …) can't be used as
  variable or parameter names — doing so causes `'assign`.

## Repo-specific gotcha: `update/delete from` a symbol

`update ... from \`t` (backtick-quoted table name) updates the global
variable **in place** and evaluates to the table's *name* (a symbol), not
its content. `update ... from t` (bare variable, no backtick) evaluates to
the updated table's *value*.

This matters because reassigning the symbol form clobbers the table:

```q
q) t:([]a:100 200)
q) t:update a:.xls.as'[`s65;a] from `t   / BUG: t now IS the symbol `t`
q) t
`t
```

The fix is to drop the reassignment — the in-place update already did the
work:

```q
q) update a:.xls.as'[`s65;a] from `t     / correct: bare statement, no `t:`
q) t                                     / t is still the real table
```

Every `.xls.as` example in [README.md](../README.md) uses the bare form for
this reason. If you see `t:update ... from \`t` anywhere, it's a bug.

## Running/testing q locally

Don't assume no `q` interpreter is available just because it's not on
`PATH`. Check, in order:

1. `type q` / `which q`, and `alias | grep -i '\bq'` — kdb+ users often
   alias versioned launchers (e.g. `q4.0`) rather than putting a bare `q`
   on `PATH`.
2. Shell rc files (`~/.bash_profile`, `~/.zshrc`, `~/.zprofile`) for
   `QHOME`/`QBASE`/`QLIC` exports — kdb+'s own convention is a per-version
   directory (e.g. `$QBASE/4.0-2022.05.11/m64/q`) plus a `k4.lic` license
   file, wired up via one of these variables rather than a symlink on
   `PATH`.
3. On Apple Silicon, a legacy kdb+ macOS build is x86_64 Mach-O, not
   arm64 — it still runs fine under Rosetta 2, no emulation needed. Confirm
   with `file <path>/q` and `arch -x86_64 /usr/bin/true`.
4. If no local binary + license turns up, a Linux `q` binary can be run via
   Docker with `--platform=linux/amd64` (Docker Desktop's VM handles the
   emulation) — mount the binary and license read-only, set `QHOME`/`QLIC`
   accordingly. Two things that silently break this: `docker run` needs
   `-i` or heredoc input to the `q` process is dropped with no error, and
   `-q` should be passed after the script filename
   (`q script.q -q <<'EOF' ... EOF`) so the script loads before stdin is
   read.
5. Always end a scripted q session with `exit 0`, or the process hangs
   waiting for more input.

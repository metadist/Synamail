/**
 * commitlint config — used by .githooks/commit-msg once node_modules exists
 * (Sprint 2.1 onwards) and by the GitHub Actions `commitlint` job.
 *
 * Mirrors docs/COMMIT_PROCESS.md §1.
 */
module.exports = {
  extends: ['@commitlint/config-conventional'],
  rules: {
    'type-enum': [
      2,
      'always',
      [
        'feat',
        'fix',
        'refactor',
        'perf',
        'test',
        'docs',
        'build',
        'ci',
        'chore',
        'style',
        'revert',
      ],
    ],
    'scope-enum': [
      1, // warn — keeps history searchable without blocking valid one-offs
      'always',
      [
        'taskpane',
        'commands',
        'dialog',
        'auth',
        'client',
        'manifest',
        'i18n',
        'gui',
        'rag',
        'rules',
        'contact-kb',
        'tests',
        'ci',
        'docs',
        'deps',
        'assets',
        'plan',
      ],
    ],
    // Allow `sentence-case` because Dependabot writes subjects like
    // "ci(deps): Bump actions/checkout from 4 to 6". Block only the louder
    // styles (Start-Case, PascalCase, ALL CAPS).
    'subject-case': [2, 'never', ['start-case', 'pascal-case', 'upper-case']],
    'subject-empty': [2, 'never'],
    'subject-full-stop': [2, 'never', '.'],
    // Dependabot's group-bump subjects can reach ~80 chars
    // ("build(deps-dev): Bump the minor-and-patch group across 1 directory with 2 updates").
    // 100 is a common upper bound that still keeps `git log --oneline` readable.
    'header-max-length': [2, 'always', 100],
    'body-leading-blank': [2, 'always'],
    'footer-leading-blank': [2, 'always'],
    // Disabled: Dependabot bodies always contain release-note URLs that exceed
    // 100 chars and can't reasonably be wrapped. Subject + header are still
    // capped at 72 by `header-max-length`.
    'body-max-line-length': [0, 'always', 100],
  },
}

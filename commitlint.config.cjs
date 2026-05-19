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
    'subject-case': [
      2,
      'never',
      ['sentence-case', 'start-case', 'pascal-case', 'upper-case'],
    ],
    'subject-empty': [2, 'never'],
    'subject-full-stop': [2, 'never', '.'],
    'header-max-length': [2, 'always', 72],
    'body-leading-blank': [2, 'always'],
    'footer-leading-blank': [2, 'always'],
  },
}

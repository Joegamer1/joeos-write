# Publication verification

This public WIP snapshot includes source, project artwork and third-party notices. Personal documents and paths, private audit materials, build outputs, local profiles, dependencies and prior Git history are excluded. Selected text files were screened for credential patterns and personal identifiers; this is not a guarantee of security or a full code audit.

29 synthetic in-memory backend tests passed across preservation, text editing, package storage and layout model tests. Two filesystem-writing tests were excluded from this run because the publication environment is read-only: `test_failed_replace_preserves_target` and `test_preview_preferences_are_independent_and_atomic`. They remain in the source for normal local testing. Full desktop/UI tests, packaging and visual equivalence with Microsoft Word were not rerun for publication.

No binaries or font files are distributed in this snapshot. Test results do not establish production readiness.

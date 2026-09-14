<?php

declare(strict_types=1);

/**
 * Cross-engine OPS parity helper: diff two documents with the PHP last-word and
 * replay the ops, so the TS port's Agent.diff / Agent.reduce can be compared op
 * for op. Uses the same minimal PSR-4 autoloader as php-tobytes.php (the PHP
 * core is zero-dependency) — no composer needed.
 *
 *   php php-diff.php <case.json>     {"a": ..., "b": ...}
 *                                    -> {"ops": Agent::diff(a, b), "reduced": Agent::reduce(a, ops)}
 *
 * The file may also hold a LIST of cases, answered with a list of results in the
 * same order, so a suite pays for one PHP process rather than one per case. A
 * case carrying "ops" instead of "b" replays those ops rather than diffing,
 * which pins the reducer on its own. A case {"hunks": [a, b]} answers
 * {"hunks": DocDiff::hunks(a, b)}, which pins the alignment and its tie-break
 * directly. A case that throws answers {"error": ...}.
 *
 *   php php-diff.php --op-schema     -> Agent::opSchema()
 *
 * `LAST_WORD_PHP_SRC` is checked before the sibling checkout, as in
 * php-tobytes.php: a hard-coded sibling path alone resolves only inside the .agi
 * envelope, so CI would silently get no parity run at all rather than an error.
 */

spl_autoload_register(function (string $class): void {
    $prefix = 'LastWord\\';
    if (strncmp($class, $prefix, strlen($prefix)) !== 0) {
        return;
    }
    $rel = substr($class, strlen($prefix));
    $root = getenv('LAST_WORD_PHP_SRC') ?: __DIR__.'/../../last-word/src';
    $file = rtrim($root, '/').'/'.str_replace('\\', '/', $rel).'.php';
    if (is_file($file)) {
        require $file;
    }
});

// Fail loudly if the autoloader found nothing, or found a PHP last-word too old
// to diff. Without this every case answers "class not found", which reads like
// a parity failure rather than a missing or stale checkout.
if (! class_exists(\LastWord\Agent::class) || ! class_exists(\LastWord\Ops\DocDiff::class)) {
    fwrite(STDERR, "LastWord\\Agent or LastWord\\Ops\\DocDiff not found. Set LAST_WORD_PHP_SRC to the src/ directory of PHP last-word 0.6.0 or later.\n");
    exit(3);
}

$flags = JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_THROW_ON_ERROR;

if (($argv[1] ?? null) === '--op-schema') {
    echo json_encode(\LastWord\Agent::opSchema(), $flags);
    exit(0);
}

if ($argc < 2) {
    fwrite(STDERR, "usage: php php-diff.php <cases.json> | --op-schema\n");
    exit(2);
}

$input = json_decode((string) file_get_contents($argv[1]), true, 512, JSON_THROW_ON_ERROR);

$run = static function (array $case): array {
    try {
        if (array_key_exists('hunks', $case)) {
            return ['hunks' => \LastWord\Ops\DocDiff::hunks($case['hunks'][0], $case['hunks'][1])];
        }

        $ops = array_key_exists('ops', $case)
            ? $case['ops']
            : \LastWord\Agent::diff($case['a'], $case['b']);

        return ['ops' => $ops, 'reduced' => \LastWord\Agent::reduce($case['a'], $ops)];
    } catch (\Throwable $e) {
        return ['error' => get_class($e).': '.$e->getMessage()];
    }
};

echo json_encode(
    array_is_list($input) ? array_map($run, $input) : $run($input),
    $flags,
);

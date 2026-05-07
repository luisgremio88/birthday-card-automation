<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

if ($_SERVER['REQUEST_METHOD'] !== 'GET') {
    jsonResponse([
        'success' => false,
        'message' => 'Metodo nao permitido.',
    ], 405);
}

jsonResponse([
    'success' => true,
    'emailDefaults' => emailDefaults(),
]);

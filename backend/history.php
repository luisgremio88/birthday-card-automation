<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

$profile = (string) ($_GET['profile'] ?? 'associado');
$config = profileConfig($profile);
$projectRoot = projectRoot();

$logFile = projectRoot() . DIRECTORY_SEPARATOR . 'logs' . DIRECTORY_SEPARATOR .
    ($profile === 'diretoria' ? 'envios_diretoria.csv' : 'envios_associado.csv');

if (!is_file($logFile)) {
    jsonResponse([
        'success' => true,
        'items' => [],
    ]);
}

$handle = fopen($logFile, 'rb');
if ($handle === false) {
    jsonResponse([
        'success' => false,
        'message' => 'Nao foi possivel abrir o historico.',
    ], 500);
}

$header = fgetcsv($handle);
$items = [];

while (($row = fgetcsv($handle)) !== false) {
    if ($header === false) {
        continue;
    }

    $entry = array_combine($header, $row);
    if ($entry === false) {
        continue;
    }

    $cardPath = (string) ($entry['arquivo_cartao'] ?? '');
    $relativePath = str_replace($projectRoot . DIRECTORY_SEPARATOR, '', $cardPath);
    $relativePath = str_replace(DIRECTORY_SEPARATOR, '/', $relativePath);
    $entry['arquivo_cartao_url'] = $relativePath !== '' ? "/projeto-aniversario/{$relativePath}" : '';

    $items[] = $entry;
}

fclose($handle);

$items = array_reverse($items);

jsonResponse([
    'success' => true,
    'profile' => $profile,
    'items' => $items,
]);

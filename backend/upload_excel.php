<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    jsonResponse([
        'success' => false,
        'message' => 'Metodo nao permitido.',
    ], 405);
}

if (!isset($_FILES['excel'])) {
    jsonResponse([
        'success' => false,
        'message' => 'Nenhum arquivo foi enviado.',
    ], 400);
}

$file = $_FILES['excel'];
$profile = (string) ($_POST['profile'] ?? 'associado');
$config = profileConfig($profile);

if (($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
    jsonResponse([
        'success' => false,
        'message' => 'Falha no upload da planilha.',
    ], 400);
}

$extension = strtolower(pathinfo((string) $file['name'], PATHINFO_EXTENSION));
if (!in_array($extension, ['xlsx', 'xls'], true)) {
    jsonResponse([
        'success' => false,
        'message' => 'Envie uma planilha .xlsx ou .xls.',
    ], 400);
}

$targetPath = projectRoot() . DIRECTORY_SEPARATOR . 'uploads' . DIRECTORY_SEPARATOR . $config['excel'];

if (!move_uploaded_file((string) $file['tmp_name'], $targetPath)) {
    jsonResponse([
        'success' => false,
        'message' => 'Nao foi possivel salvar a planilha em uploads/.',
    ], 500);
}

jsonResponse([
    'success' => true,
    'message' => sprintf('Planilha do perfil %s salva no servidor com sucesso.', $profile),
    'path' => $targetPath,
]);

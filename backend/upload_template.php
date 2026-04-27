<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    jsonResponse([
        'success' => false,
        'message' => 'Metodo nao permitido.',
    ], 405);
}

if (!isset($_FILES['templatePsd'])) {
    jsonResponse([
        'success' => false,
        'message' => 'Nenhum PSD foi enviado.',
    ], 400);
}

$file = $_FILES['templatePsd'];
$profile = (string) ($_POST['profile'] ?? 'associado');
$config = profileConfig($profile);

if (($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
    jsonResponse([
        'success' => false,
        'message' => 'Falha no upload do PSD.',
    ], 400);
}

$extension = strtolower(pathinfo((string) $file['name'], PATHINFO_EXTENSION));
if ($extension !== 'psd') {
    jsonResponse([
        'success' => false,
        'message' => 'Envie um arquivo PSD do Photoshop.',
    ], 400);
}

$templateDir = projectRoot() . DIRECTORY_SEPARATOR . 'templates';
if (!is_dir($templateDir) && !mkdir($templateDir, 0777, true) && !is_dir($templateDir)) {
    jsonResponse([
        'success' => false,
        'message' => 'Nao foi possivel preparar a pasta de templates.',
    ], 500);
}

$targetPath = $templateDir . DIRECTORY_SEPARATOR . $config['templatePsd'];

if (!move_uploaded_file((string) $file['tmp_name'], $targetPath)) {
    jsonResponse([
        'success' => false,
        'message' => 'Nao foi possivel salvar o PSD em templates/.',
    ], 500);
}

$scriptPath = projectRoot() . DIRECTORY_SEPARATOR . 'automacao' . DIRECTORY_SEPARATOR . 'extrair_template_limpo.py';
$command = sprintf(
    'python %s --profile %s --psd %s 2>&1',
    escapeshellarg($scriptPath),
    escapeshellarg($profile),
    escapeshellarg($targetPath)
);

$output = [];
$exitCode = 0;
exec($command, $output, $exitCode);

if ($exitCode !== 0) {
    jsonResponse([
        'success' => false,
        'message' => 'O PSD foi salvo, mas a extracao da base limpa falhou.',
        'output' => $output,
    ], 500);
}

jsonResponse([
    'success' => true,
    'message' => sprintf('PSD do perfil %s enviado com sucesso e base limpa regenerada.', $profile),
    'path' => $targetPath,
    'output' => $output,
]);

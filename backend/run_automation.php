<?php

declare(strict_types=1);

require __DIR__ . '/bootstrap.php';

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    jsonResponse([
        'success' => false,
        'message' => 'Metodo nao permitido.',
    ], 405);
}

$rawInput = file_get_contents('php://input');
$payload = json_decode($rawInput ?: '{}', true);

$profile = (string) ($payload['profile'] ?? 'associado');
$config = profileConfig($profile);
$mode = $payload['mode'] ?? 'draft';
$senderEmail = trim((string) ($payload['senderEmail'] ?? 'luis.dias@cnbrs.org.br'));

$scriptPath = projectRoot() . DIRECTORY_SEPARATOR . 'automacao' . DIRECTORY_SEPARATOR . 'enviar_aniversarios.py';

if (!is_file($scriptPath)) {
    jsonResponse([
        'success' => false,
        'message' => 'Script de automacao nao encontrado.',
    ], 500);
}

$commandParts = [
    'python',
    escapeshellarg($scriptPath),
    '--profile',
    escapeshellarg($profile),
    '--sender-email',
    escapeshellarg($senderEmail),
];

if ($mode === 'send') {
    $commandParts[] = '--send';
}

$command = implode(' ', $commandParts) . ' 2>&1';
$output = [];
$exitCode = 0;
exec($command, $output, $exitCode);

$joinedOutput = trim(implode(PHP_EOL, $output));

if ($exitCode !== 0) {
    jsonResponse([
        'success' => false,
        'message' => $joinedOutput ?: 'A automacao retornou erro.',
        'output' => $output,
    ], 500);
}

$message = $mode === 'send'
    ? sprintf('Automacao do perfil %s executada com envio real. Verifique o Outlook e o log.', $profile)
    : sprintf('Automacao do perfil %s executada em modo rascunho. O Outlook deve abrir os e-mails para conferencia.', $profile);

jsonResponse([
    'success' => true,
    'message' => $message,
    'output' => $output,
]);

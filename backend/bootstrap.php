<?php

declare(strict_types=1);

header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

function jsonResponse(array $payload, int $statusCode = 200): void
{
    http_response_code($statusCode);
    header('Content-Type: application/json; charset=utf-8');
    $encoded = json_encode(normalizeJsonPayload($payload), JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    if ($encoded === false) {
        $encoded = json_encode([
            'success' => false,
            'message' => 'Falha ao montar resposta JSON: ' . json_last_error_msg(),
        ], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    }

    echo $encoded ?: '{"success":false,"message":"Falha ao montar resposta JSON."}';
    exit;
}

function normalizeJsonPayload($value)
{
    if (is_array($value)) {
        return array_map('normalizeJsonPayload', $value);
    }

    if (!is_string($value) || preg_match('//u', $value)) {
        return $value;
    }

    $converted = @iconv('Windows-1252', 'UTF-8//IGNORE', $value);
    if ($converted !== false) {
        return $converted;
    }

    return utf8_encode($value);
}

function projectRoot(): string
{
    return dirname(__DIR__);
}

function emailDefaults(): array
{
    $defaults = [
        'senderEmail' => 'aniversarios@exemplo.com',
        'bccEmail' => 'auditoria@exemplo.com',
    ];
    $configPath = projectRoot() . DIRECTORY_SEPARATOR . 'config' . DIRECTORY_SEPARATOR . 'email_defaults.json';

    if (!is_file($configPath)) {
        return $defaults;
    }

    $content = file_get_contents($configPath);
    $config = json_decode($content ?: '{}', true);
    if (!is_array($config)) {
        return $defaults;
    }

    return [
        'senderEmail' => trim((string) ($config['senderEmail'] ?? $defaults['senderEmail'])),
        'bccEmail' => trim((string) ($config['bccEmail'] ?? $defaults['bccEmail'])),
    ];
}

function profileConfig(string $profile): array
{
    $profiles = [
        'associado' => [
            'excel' => 'aniversariantes_associado.xlsx',
            'templatePsd' => 'cartao_template_associado.psd',
        ],
        'diretoria' => [
            'excel' => 'aniversariantes_diretoria.xlsx',
            'templatePsd' => 'cartao_template_diretoria.psd',
        ],
    ];

    if (!isset($profiles[$profile])) {
        jsonResponse([
            'success' => false,
            'message' => 'Perfil invalido.',
        ], 400);
    }

    return $profiles[$profile];
}

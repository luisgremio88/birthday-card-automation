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
    echo json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

function projectRoot(): string
{
    return dirname(__DIR__);
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

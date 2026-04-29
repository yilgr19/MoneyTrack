<?php

return [
    /*
    |--------------------------------------------------------------------------
    | Clave compartida app ↔ API de registro de tokens
    |--------------------------------------------------------------------------
    | La app envía la misma clave en la cabecera X-MoneyTrack-Api-Key.
    | En producción debe estar definida.
    */
    'push_api_key' => env('MONEYTRACK_PUSH_API_KEY'),

    /*
    | Expo Push API (no suele necesitar cambio).
    */
    'expo_push_url' => env('MONEYTRACK_EXPO_PUSH_URL', 'https://exp.host/--/api/v2/push/send'),
];

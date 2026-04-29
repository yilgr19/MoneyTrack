# MoneyTrack — registro de push (Laravel) + Expo

Fragmentos listos para integrar en un proyecto **Laravel 10+** (también válido en Laravel 11).

## 1. Copiar archivos

- `app/Models/MoneyTrackExpoPushToken.php`
- `app/Http/Controllers/Api/MoneyTrackExpoPushTokenController.php`
- `app/Http/Requests/RegisterMoneyTrackExpoPushTokenRequest.php`
- `app/Http/Middleware/ValidateMoneyTrackPushApiKey.php`
- `app/Services/ExpoPushService.php`
- `app/Jobs/SendMoneyTrackExpoPushJob.php`
- `config/moneytrack.php`
- `database/migrations/2026_04_28_120000_create_moneytrack_expo_push_tokens_table.php`

## 2. Variables de entorno (`.env`)

```env
MONEYTRACK_PUSH_API_KEY=cambia-por-un-secreto-largo-aleatorio
```

Si dejas la clave vacía en **local**, el middleware no exige cabecera (útil para pruebas). En **production** debe estar definida o las rutas rechazan el registro.

## 3. Rutas y middleware

**Laravel 11:** en `bootstrap/app.php`, registra el alias:

```php
->withMiddleware(function (Middleware $middleware) {
    $middleware->alias([
        'moneytrack.push' => \App\Http\Middleware\ValidateMoneyTrackPushApiKey::class,
    ]);
})
```

**Laravel 10:** en `app/Http/Kernel.php`, dentro de `$middlewareAliases` (o `$routeMiddleware`):

```php
'moneytrack.push' => \App\Http\Middleware\ValidateMoneyTrackPushApiKey::class,
```

En `routes/api.php`:

```php
Route::prefix('v1/moneytrack')
    ->middleware(['throttle:60,1', 'moneytrack.push'])
    ->group(function () {
        Route::post('push-tokens', [\App\Http\Controllers\Api\MoneyTrackExpoPushTokenController::class, 'store']);
    });
```

Ajusta el prefijo o el throttle según tu política.

## 4. Cola (envío diferido)

El job `SendMoneyTrackExpoPushJob` implementa `ShouldQueue`. Configura `QUEUE_CONNECTION` (por ejemplo `redis` o `database` en producción).

## 5. App móvil

Define en el entorno de build de Expo:

- `EXPO_PUBLIC_PUSH_REGISTER_URL` — URL completa del endpoint, p. ej. `https://tu-dominio.com/api/v1/moneytrack/push-tokens`
- `EXPO_PUBLIC_PUSH_API_KEY` — mismo valor que `MONEYTRACK_PUSH_API_KEY` (solo protege frente a registro masivo anónimo; para máxima seguridad usa además usuarios autenticados y guarda el token por `user_id`).

## 6. Ejemplo: disparar una notificación desde código

```php
use App\Jobs\SendMoneyTrackExpoPushJob;
use App\Models\MoneyTrackExpoPushToken;

$tokens = MoneyTrackExpoPushToken::query()->pluck('expo_push_token')->all();

SendMoneyTrackExpoPushJob::dispatch(
    $tokens,
    'Recordatorio',
    'Tienes un pago programado próximo.',
    [
        'rootScreen' => 'Mas',
        'nestedScreen' => 'PagosProgramados',
        'nestedParams' => new \stdClass(),
    ]
);
```

Los datos en el último argumento deben ser JSON-serializables (objetos vacíos como `{}` en la app: usa `new \stdClass()` o `[]` según necesidad de Expo).

<?php

namespace App\Http\Middleware;

use Closure;
use Illuminate\Http\Request;
use Symfony\Component\HttpFoundation\Response;

class ValidateMoneyTrackPushApiKey
{
    public function handle(Request $request, Closure $next): Response
    {
        $expected = config('moneytrack.push_api_key');

        if (app()->environment('production') && (! is_string($expected) || $expected === '')) {
            abort(503, 'MoneyTrack push API is not configured.');
        }

        if (! is_string($expected) || $expected === '') {
            return $next($request);
        }

        $given = $request->header('X-MoneyTrack-Api-Key', '');

        if (! hash_equals($expected, (string) $given)) {
            abort(401, 'Invalid API key.');
        }

        return $next($request);
    }
}

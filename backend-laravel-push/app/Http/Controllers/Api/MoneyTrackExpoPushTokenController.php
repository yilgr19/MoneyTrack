<?php

namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use App\Http\Requests\RegisterMoneyTrackExpoPushTokenRequest;
use App\Models\MoneyTrackExpoPushToken;
use Illuminate\Http\JsonResponse;

class MoneyTrackExpoPushTokenController extends Controller
{
    public function store(RegisterMoneyTrackExpoPushTokenRequest $request): JsonResponse
    {
        $validated = $request->validated();

        $row = MoneyTrackExpoPushToken::query()->updateOrCreate(
            ['device_install_id' => $validated['deviceInstallId']],
            [
                'expo_push_token' => $validated['expoPushToken'],
                'platform' => $validated['platform'],
                'last_registered_at' => now(),
            ]
        );

        return response()->json([
            'ok' => true,
            'id' => $row->id,
        ]);
    }
}

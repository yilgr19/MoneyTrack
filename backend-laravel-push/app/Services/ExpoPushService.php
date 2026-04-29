<?php

namespace App\Services;

use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class ExpoPushService
{
    /**
     * @param  array<int, array<string, mixed>>  $messages  Cada elemento: to, title, body, data?, sound?, channelId?, priority?
     * @return array<string, mixed>
     */
    public function sendMessages(array $messages): array
    {
        if ($messages === []) {
            return ['data' => []];
        }

        $url = (string) config('moneytrack.expo_push_url');

        $response = Http::acceptJson()
            ->timeout(15)
            ->post($url, ['messages' => $messages]);

        if (! $response->successful()) {
            Log::warning('Expo push request failed', [
                'status' => $response->status(),
                'body' => $response->body(),
            ]);
        }

        return $response->json() ?? [];
    }
}

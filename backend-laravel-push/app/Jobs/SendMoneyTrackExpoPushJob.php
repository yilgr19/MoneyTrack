<?php

namespace App\Jobs;

use App\Services\ExpoPushService;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;

class SendMoneyTrackExpoPushJob implements ShouldQueue
{
    use Dispatchable;
    use InteractsWithQueue;
    use Queueable;
    use SerializesModels;

    /**
     * @param  array<int, string>  $expoPushTokens
     * @param  array<string, mixed>  $data  Payload en `data` de la notificación (deep link en la app).
     */
    public function __construct(
        public array $expoPushTokens,
        public string $title,
        public string $body,
        public array $data = [],
    ) {}

    public function handle(ExpoPushService $expoPush): void
    {
        $tokens = array_values(array_filter(array_unique($this->expoPushTokens)));
        if ($tokens === []) {
            return;
        }

        $messages = [];
        foreach ($tokens as $to) {
            $messages[] = [
                'to' => $to,
                'title' => $this->title,
                'body' => $this->body,
                'sound' => 'default',
                'channelId' => 'default',
                'priority' => 'high',
                'data' => $this->data,
            ];
        }

        $expoPush->sendMessages($messages);
    }
}

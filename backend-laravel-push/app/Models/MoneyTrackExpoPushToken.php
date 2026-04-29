<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\BelongsTo;

class MoneyTrackExpoPushToken extends Model
{
    protected $table = 'moneytrack_expo_push_tokens';

    protected $fillable = [
        'device_install_id',
        'expo_push_token',
        'platform',
        'user_id',
        'last_registered_at',
    ];

    protected $casts = [
        'last_registered_at' => 'datetime',
    ];

    public function user(): BelongsTo
    {
        return $this->belongsTo(User::class);
    }
}

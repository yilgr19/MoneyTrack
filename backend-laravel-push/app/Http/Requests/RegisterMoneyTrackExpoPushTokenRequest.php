<?php

namespace App\Http\Requests;

use Illuminate\Foundation\Http\FormRequest;

class RegisterMoneyTrackExpoPushTokenRequest extends FormRequest
{
    public function authorize(): bool
    {
        return true;
    }

    public function rules(): array
    {
        return [
            'expoPushToken' => ['required', 'string', 'max:512'],
            'platform' => ['required', 'string', 'in:ios,android'],
            'deviceInstallId' => ['required', 'string', 'max:64', 'regex:/^[a-zA-Z0-9\-_.]+$/'],
        ];
    }
}

<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    public function up(): void
    {
        Schema::create('moneytrack_expo_push_tokens', function (Blueprint $table) {
            $table->id();
            $table->string('device_install_id', 64)->unique();
            $table->string('expo_push_token', 512)->index();
            $table->string('platform', 16);
            $table->unsignedBigInteger('user_id')->nullable()->index();
            $table->timestamp('last_registered_at')->nullable();
            $table->timestamps();
        });
    }

    public function down(): void
    {
        Schema::dropIfExists('moneytrack_expo_push_tokens');
    }
};

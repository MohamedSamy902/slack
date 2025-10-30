<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\SlackDataController;

Route::get('/', function () {
    return view('welcome');
});

Route::get('/report/export', [SlackDataController::class, 'exportReport'])->name('slack.report.export');


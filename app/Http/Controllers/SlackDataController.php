<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Http;
use Maatwebsite\Excel\Facades\Excel;
use App\Exports\SlackMonthlyReportExport;
use Carbon\Carbon;
use Carbon\CarbonPeriod;

class SlackDataController extends Controller
{
    // *** يرجى تغيير هذا الثابت إلى المنطقة الزمنية المطلوبة ***
    private const TARGET_TIMEZONE = 'Africa/Cairo'; // مثال: توقيت القاهرة
    // ************************************************************
    // *** تعريف ساعات العمل القياسية (8 ساعات) ***
    private const STANDARD_WORK_HOURS = 8 * 3600; // 8 ساعات بالثواني
    // *** تعريف نقطة الفصل بين الدخول والخروج (16:00) ***
    private const CHECK_OUT_HOUR_THRESHOLD = 16;

    /**
     * يستخرج بيانات الرسائل من قناة Slack محددة ويقوم بتجميعها.
     *
     * @param Request $request
     * @return array مصفوفة البيانات المجمعة (ليست استجابة JSON).
     */
    private function getRawChannelMessages(Request $request): array
    {
        // 1. المتغيرات المطلوبة
        $token = env('SLACK_BOT_TOKEN');
        $channelId = env('SLACK_CHANNEL_ID');
        $cursor = $request->get('cursor');

        // تحديد المنطقة الزمنية لمعالجة الطوابع الزمنية بشكل صحيح
        date_default_timezone_set(self::TARGET_TIMEZONE);

        if (!$token || !$channelId) {
            // في بيئة الإنتاج، يجب التعامل مع هذا بشكل أفضل
            throw new \Exception('Slack token or Channel ID not set in .env file.');
        }

        // --- حساب الطابع الزمني لليوم الأول من الشهر الحالي ---
        $firstDayOfMonth = date('Y-m-01 00:00:00');
        $oldestTimestamp = strtotime($firstDayOfMonth);
        // --------------------------------------------------------

        $rawMessages = [];
        $nextCursor = $cursor;
        $maxPages = 5;

        // 2. حلقة لجلب سجل الرسائل بالترحيل (Pagination)
        do {
            $response = Http::withToken($token)
                ->get('https://slack.com/api/conversations.history', [
                    'channel' => $channelId,
                    'limit' => 200,
                    'cursor' => $nextCursor,
                    'oldest' => $oldestTimestamp,
                ]);

            $data = $response->json();
            if (!$data['ok']) {
                throw new \Exception('Slack API Error: ' . ($data['error'] ?? 'Unknown error'));
            }

            // 3. تحليل وتصفية الرسائل التي أرسلها المستخدمون فقط
            foreach ($data['messages'] as $message) {
                // تصفية الرسائل التي أرسلها المستخدمون ولا تحتوي على subtypes (مثل رسائل النظام)
                if (isset($message['user']) && !isset($message['subtype'])) {
                    $rawMessages[] = [
                        'user_id' => $message['user'],
                        'timestamp' => (float) $message['ts'],
                        'text' => $message['text'] ?? '',
                    ];
                }
            }

            $nextCursor = $data['response_metadata']['next_cursor'] ?? null;
            $maxPages--;

            if ($nextCursor) {
                // تأخير لتجنب تخطي حدود حدود معدل Slack
                usleep(500000); // 0.5 ثانية
            }
        } while ($nextCursor && $maxPages > 0);

        return $rawMessages;
    }


    /**
     * تجميع الرسائل حسب المستخدم واليوم وتحديد الدخول/الخروج بناءً على الساعة 16:00.
     *
     * @param array $messages مصفوفة الرسائل الخام.
     * @return array مصفوفة مجمعة حسب المستخدم والتاريخ.
     */
    private function groupMessagesByDayAndUser(array $messages): array
    {
        $report = [];

        // فرز الرسائل حسب الطابع الزمني لضمان تحديد الدخول والخروج بشكل صحيح
        usort($messages, function ($a, $b) {
            return $a['timestamp'] <=> $b['timestamp'];
        });

        foreach ($messages as $message) {
            $userId = $message['user_id'];
            $timestamp = $message['timestamp'];

            $dateTime = Carbon::createFromTimestamp((int) $timestamp, self::TARGET_TIMEZONE);
            $date = $dateTime->format('Y-m-d');

            if (!isset($report[$userId][$date])) {
                $report[$userId][$date] = [
                    'date' => $date,
                    'check_ins' => [], // رسائل قبل 16:00
                    'check_outs' => [], // رسائل بعد 16:00
                    'messages' => [],
                ];
            }

            $messageData = [
                // *** تم التعديل هنا: استخدام تنسيق 12 ساعة (h:i:s A) بدلاً من 24 ساعة (H:i:s) ***
                'time' => $dateTime->format('h:i:s A'),
                // ******************************************************************************
                'timestamp' => $timestamp,
                'text' => $message['text'],
            ];

            // تجميع الرسائل حسب نقطة الفصل 16:00 (4:00 PM)
            if ($dateTime->hour < self::CHECK_OUT_HOUR_THRESHOLD) {
                $report[$userId][$date]['check_ins'][] = $messageData;
            } else {
                $report[$userId][$date]['check_outs'][] = $messageData;
            }

            $report[$userId][$date]['messages'][] = $messageData;
        }

        $finalGrouped = [];
        foreach ($report as $userId => $days) {
            foreach ($days as $data) {

                // 1. تحديد الدخول (أقدم رسالة قبل 16:00)
                $firstCheckIn = null;
                if (!empty($data['check_ins'])) {
                    $firstCheckIn = reset($data['check_ins']);
                }

                // 2. تحديد الخروج (أحدث رسالة بعد 16:00)
                $lastCheckOut = null;
                if (!empty($data['check_outs'])) {
                    $lastCheckOut = end($data['check_outs']);
                }

                // 3. تحديد أحدث رسالة على الإطلاق (لإجمالي الساعات)
                $firstMessage = reset($data['messages']);
                $lastMessage = end($data['messages']);

                // تحديد أوقات الدخول/الخروج المنسقة
                $inTime = $firstCheckIn['time'] ?? 'N/A';
                $outTime = $lastCheckOut['time'] ?? 'N/A';

                $finalGrouped[] = [
                    'user_id' => $userId,
                    'date' => $data['date'],
                    // الدخول والخروج الفعليين (بناءً على الساعة 16:00)
                    'first_check_in' => $inTime,
                    'last_check_out' => $outTime,
                    // الطوابع الزمنية المستخدمة لحساب إجمالي الساعات (أقدم وأحدث رسالة بغض النظر عن 16:00)
                    'first_timestamp' => $firstMessage['timestamp'] ?? null,
                    'last_timestamp' => $lastMessage['timestamp'] ?? null,
                    'all_messages' => $data['messages'],
                    'message_count' => count($data['messages']),
                    // حالة التسجيل بناءً على المنطق الجديد
                    'has_check_in' => !is_null($firstCheckIn),
                    'has_check_out' => !is_null($lastCheckOut),
                ];
            }
        }

        // فرز التقرير النهائي حسب التاريخ
        usort($finalGrouped, function ($a, $b) {
            return strtotime($a['date']) - strtotime($b['date']);
        });

        return $finalGrouped;
    }


    /**
     * يحوّل معرّفات المستخدمين إلى أسماء قابلة للقراءة ويضيف الحسابات الأساسية.
     *
     * @param array $groupedEvents مصفوفة الأحداث المجمعة.
     * @param string $token رمز الوصول.
     * @return array التقرير مجمّع حسب اسم المستخدم ويحتوي على البيانات اليومية المحسوبة.
     */
    private function resolveUserNamesAndFormatReport(array $groupedEvents, string $token): array
    {
        $uniqueUserIds = array_unique(array_column($groupedEvents, 'user_id'));
        $userNameCache = [];

        // 1. جلب أسماء المستخدمين وتخزينها مؤقتاً (كما في السابق)
        // ... (جزء جلب الأسماء كما هو)
        //     // 1. جلب أسماء المستخدمين وتخزينها مؤقتاً (كما في السابق)
        foreach ($uniqueUserIds as $userId) {
            $response = Http::withToken($token)
                ->get('https://slack.com/api/users.info', ['user' => $userId]);

            $userData = $response->json();
            if ($userData['ok'] && isset($userData['user']['real_name'])) {
                $userNameCache[$userId] = $userData['user']['real_name'];
            } else {
                $userNameCache[$userId] = 'مستخدم غير معروف (' . $userId . ')';
            }
            usleep(200000);
        }
        // 2. تجميع البيانات حسب اسم المستخدم وإضافة الحسابات
        $reportByUser = [];
        foreach ($groupedEvents as $event) {
            $userName = $userNameCache[$event['user_id']] ?? $event['user_id'];

            // حساب إجمالي الساعات بناءً على أول وآخر رسالة على الإطلاق (لإجمالي ساعات التواجد)
            $totalSeconds = 0;
            if ($event['first_timestamp'] && $event['last_timestamp']) {
                $totalSeconds = $event['last_timestamp'] - $event['first_timestamp'];
            }

            // حساب فرق الوقت
            $timeDifferenceSeconds = $totalSeconds - self::STANDARD_WORK_HOURS;

            // تنسيق الوقت
            $totalHoursFormatted = gmdate('H:i:s', abs($totalSeconds));
            $timeDifferenceFormatted = ($timeDifferenceSeconds < 0 ? '-' : '') . gmdate('H:i:s', abs($timeDifferenceSeconds));

            // حالة الرسائل لورقة العمل:
            $status = 'كامل';

            // تحديد حالة التسجيل الناقص بناءً على المنطق الجديد
            if (!$event['has_check_in'] && !$event['has_check_out']) {
                $status = 'غائب'; // لم يتم العثور على دخول ولا خروج
            } elseif (!$event['has_check_in'] || !$event['has_check_out']) {
                $status = 'تسجيل ناقص'; // تم العثور على دخول أو خروج فقط
            }

            // ******************************************************************************
            // *** بداية التعديل المطلوب: البحث عن 'permission' أو 'Pr' للتصريح ***
            // ******************************************************************************

            // نبحث عن التصريح في حالتين: 1) إذا كان هناك تسجيل ناقص، 2) إذا كان هناك تأخير (ساعات العمل أقل من 8)
            if ($status === 'تسجيل ناقص' || $timeDifferenceSeconds < 0) {
                $isPermissionGranted = false;

                // تجميع كل نصوص الرسائل في سلسلة واحدة للبحث
                $allTexts = implode(' ', array_column($event['all_messages'], 'text'));

                // تحويل النص إلى حروف صغيرة (Lowercase) للبحث بدون حساسية لحالة الأحرف
                $lowerCaseText = strtolower($allTexts);

                // البحث عن 'pr' أو 'permission'
                if (str_contains($lowerCaseText, 'pr') || str_contains($lowerCaseText, 'permission')) {
                    $isPermissionGranted = true;
                }

                if ($isPermissionGranted) {
                    $status = 'تصريح'; // تم تغيير الحالة إلى 'تصريح'
                }
            }
            // ******************************************************************************
            // *** نهاية التعديل المطلوب ***
            // ******************************************************************************


            $dailyData = [
                'user_id' => $event['user_id'],
                'date' => $event['date'],
                'first_check_in' => $event['first_check_in'], // الدخول قبل 16:00
                'last_check_out' => $event['last_check_out'], // الخروج بعد 16:00
                'total_hours_seconds' => $totalSeconds,
                'total_hours_formatted' => $totalHoursFormatted,
                'time_difference_seconds' => $timeDifferenceSeconds,
                'time_difference_formatted' => $timeDifferenceFormatted,
                'status' => $status,
                // تمرير حالة الدخول/الخروج الجديدة لتطبيق التنسيق في Excel
                'has_check_in' => $event['has_check_in'],
                'has_check_out' => $event['has_check_out'],
                'all_messages' => $event['all_messages'],
            ];

            $reportByUser[$userName][] = $dailyData;
        }

        return $reportByUser;
    }

    /**
     * يملأ التقرير بالتواريخ المفقودة لهذا الشهر لإنشاء سجل شهري كامل.
     *
     * @param array $reportByUser التقرير المجمع حسب المستخدم والتاريخ الفعلي للرسائل.
     * @return array التقرير الشهري الكامل لكل مستخدم.
     */
    private function generateFullMonthlyReport(array $reportByUser): array
    {
        $currentDate = Carbon::now(self::TARGET_TIMEZONE);
        $startDate = $currentDate->copy()->startOfMonth();
        $endDate = $currentDate->copy()->endOfMonth();

        $fullMonthlyReport = [];

        // إنشاء فترة زمنية تشمل جميع أيام الشهر الحالي
        $period = CarbonPeriod::create($startDate, $endDate);

        // هيكلة البيانات للمستخدمين
        foreach ($reportByUser as $userName => $dailyEvents) {
            $indexedEvents = collect($dailyEvents)->keyBy('date')->all();
            $fullMonthlyReport[$userName] = [];

            foreach ($period as $date) {
                $dateString = $date->format('Y-m-d');
                // اسم اليوم باللغة العربية (يتطلب تهيئة اللغة العربية لـ Carbon في دالة exportReport)
                $dayName = $date->translatedFormat('l');

                if (isset($indexedEvents[$dateString])) {
                    // البيانات موجودة
                    $fullMonthlyReport[$userName][] = array_merge($indexedEvents[$dateString], ['day_name' => $dayName]);
                } else {
                    // البيانات مفقودة (غائب)
                    $fullMonthlyReport[$userName][] = [
                        'user_id' => $dailyEvents[0]['user_id'] ?? 'N/A', // أخذ ID من أي يوم سابق
                        'date' => $dateString,
                        'day_name' => $dayName,
                        'first_check_in' => 'N/A',
                        'last_check_out' => 'N/A',
                        'total_hours_seconds' => 0,
                        'total_hours_formatted' => '00:00:00',
                        'time_difference_seconds' => -self::STANDARD_WORK_HOURS, // -8 ساعات
                        'time_difference_formatted' => '-08:00:00',
                        'status' => 'غائب',
                        'has_check_in' => false,
                        'has_check_out' => false,
                        'all_messages' => [],
                    ];
                }
            }
        }

        return $fullMonthlyReport;
    }


    /**
     * نقطة الدخول لإنشاء وتصدير ملف Excel متعدد الأوراق.
     * * @param Request $request
     * @return \Symfony\Component\HttpFoundation\BinaryFileResponse|\Illuminate\Http\JsonResponse
     */
    public function exportReport(Request $request)
    {
        try {
            // قم بتعيين اللغة العربية لـ Carbon للحصول على أسماء الأيام باللغة العربية
            Carbon::setLocale('ar');

            // 1. جلب الرسائل الأولية من Slack
            $rawMessages = $this->getRawChannelMessages($request);

            // 2. تجميع الرسائل حسب المستخدم والتاريخ
            $groupedEvents = $this->groupMessagesByDayAndUser($rawMessages);

            // 3. تحليل الأسماء وإضافة الحسابات الأساسية (إجمالي الساعات/الفرق)
            $reportByUserAndDay = $this->resolveUserNamesAndFormatReport($groupedEvents, env('SLACK_BOT_TOKEN'));

            // 4. إنشاء تقرير شهري كامل (بما في ذلك الأيام التي لا تحتوي على رسائل)
            $fullMonthlyReport = $this->generateFullMonthlyReport($reportByUserAndDay);

            // 5. تصدير ملف Excel
            $fileName = 'Slack_Monthly_Report_' . date('Y_m') . '.xlsx';

            // يتم تمرير البيانات المجمعة حسب المستخدم إلى فئة التصدير المتعددة الأوراق
            return Excel::download(new SlackMonthlyReportExport($fullMonthlyReport), $fileName);
        } catch (\Exception $e) {
            // يجب تسجيل الخطأ في السجلات (logs) في بيئة الإنتاج
            return response()->json([
                'error' => 'حدث خطأ أثناء إنشاء التقرير: ' . $e->getMessage()
            ], 500);
        }
    }
}

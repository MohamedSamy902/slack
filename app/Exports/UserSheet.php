<?php

namespace App\Exports;

use Carbon\Carbon;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;

class UserSheet implements FromCollection, WithTitle, WithHeadings, WithMapping, WithEvents
{
    private $userName;
    private $dailyData;

    public function __construct(string $userName, array $dailyData)
    {
        $this->userName = $userName;
        $this->dailyData = collect($dailyData);
    }

    /**
     * @return \Illuminate\Support\Collection
     */
    public function collection()
    {
        return $this->dailyData;
    }

    /**
     * @return string
     */
    public function title(): string
    {
        // اسم الشيت هو اسم المستخدم
        return $this->userName;
    }

    /**
     * @return array
     */
    public function headings(): array
    {
        // رؤوس الأعمدة باللغة العربية
        return [
            'التاريخ',
            'اليوم',
            'دخول (in)',
            'خروج (out)',
            'إجمالي الساعات',
            'فرق الوقت (Overtime/Undertime)',
            'حالة الرسائل', // عمود إضافي للمساعدة في التتبع
        ];
    }

    /**
     * تحويل بيانات الصف لتناسب شكل ورقة العمل.
     *
     * @param mixed $row
     * @return array
     */
    public function map($row): array
    {
        // يتم استخدام البيانات المُعدة مسبقًا من المتحكم
        return [
            $row['date'],
            $row['day_name'],
            $row['first_check_in'],
            $row['last_check_out'],
            $row['total_hours_formatted'],
            $row['time_difference_formatted'],
            $row['status'],
        ];
    }

    /**
     * لتطبيق التنسيق الشرطي (الألوان) على الخلايا.
     *
     * @return array
     */
    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                // إجمالي عدد الصفوف في البيانات (بالإضافة إلى صف العناوين)
                $dataRows = $this->dailyData->count();

                // التنسيق يبدأ من الصف الثاني (A2)
                for ($i = 2; $i <= $dataRows + 1; $i++) {

                    // استخراج البيانات لصف محدد (نقص 2 لأن الصفوف تبدأ من 1 والعناوين في الصف 1)
                    $rowData = $this->dailyData[$i - 2];
                    $date = Carbon::parse($rowData['date']);

                    // 1. تلوين أيام الجمعة والسبت باللون الأصفر (إجازة)
                    // 5 = الجمعة, 6 = السبت (في Carbon)
                    if ($date->dayOfWeek === Carbon::FRIDAY || $date->dayOfWeek === Carbon::SATURDAY) {
                        $event->sheet->getStyle('A' . $i . ':G' . $i)->applyFromArray([
                            'fill' => [
                                'fillType' => Fill::FILL_SOLID,
                                'color' => ['argb' => 'FFFFFF00'], // أصفر
                            ],
                        ]);
                    }

                    // 2. تلوين حالات التسجيل الناقصة باللون الأحمر
                    // التسجيل ناقص إذا كان هناك دخول فقط أو خروج فقط (وليس كلاهما 'N/A')
                    $isCheckInMissing = $rowData['first_check_in'] === 'N/A';
                    $isCheckOutMissing = $rowData['last_check_out'] === 'N/A';




                    // الحالة: تسجيل ناقص (وجود دخول فقط أو خروج فقط)
                    if (($isCheckInMissing && !$isCheckOutMissing) || (!$isCheckInMissing && $isCheckOutMissing)) {

                        // تطبيق اللون الأحمر الفاتح على الصف
                        $event->sheet->getStyle('A' . $i . ':G' . $i)->applyFromArray([
                            'fill' => [
                                'fillType' => Fill::FILL_SOLID,
                                'color' => ['argb' => 'FFFFC7CE'], // أحمر فاتح
                            ],
                        ]);
                        // تلوين خانات الدخول/الخروج المفقودة تحديداً باللون الأحمر الغامق
                        if ($isCheckInMissing) {
                            $event->sheet->getStyle('C' . $i)->applyFromArray([
                                'fill' => ['fillType' => Fill::FILL_SOLID, 'color' => ['argb' => 'FFFF9999']], // أحمر غامق للدخول المفقود
                            ]);
                        }
                        if ($isCheckOutMissing) {
                            $event->sheet->getStyle('D' . $i)->applyFromArray([
                                'fill' => ['fillType' => Fill::FILL_SOLID, 'color' => ['argb' => 'FFFF9999']], // أحمر غامق للخروج المفقود
                            ]);
                        }
                    }

                    if ($rowData['status'] == 'تصريح') {
                        // تطبيق اللون الأحمر الفاتح على الصف
                        $event->sheet->getStyle('A' . $i . ':G' . $i)->applyFromArray([
                            'fill' => [
                                'fillType' => Fill::FILL_SOLID,
                                'color' => ['argb' => 'FFFFA500'], //  برتقاني
                            ],
                        ]);
                    }
                }


                // ضبط عرض الأعمدة تلقائياً
                foreach (range('A', 'G') as $col) {
                    $event->sheet->getColumnDimension($col)->setAutoSize(true);
                }

                // تجميد الصف الأول (العناوين)
                $event->sheet->freezePane('A2');

                // تلوين صف العناوين باللون الرمادي الفاتح والخط الغامق
                $event->sheet->getStyle('A1:G1')->applyFromArray([
                    'font' => [
                        'bold' => true,
                    ],
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'color' => ['argb' => 'FFE0E0E0'],
                    ]
                ]);
            },
        ];
    }
}

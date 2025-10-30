<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

class SlackMonthlyReportExport implements WithMultipleSheets
{
    use Exportable;

    private $fullReportData;

    public function __construct(array $fullReportData)
    {
        // fullReportData هو مصفوفة مرتبة بمفتاح: اسم المستخدم والقيمة: بياناته اليومية
        $this->fullReportData = $fullReportData;
    }

    /**
     * تُرجع مصفوفة من كائنات Export/Sheet لكل ورقة عمل.
     *
     * @return array
     */
    public function sheets(): array
    {
        $sheets = [];

        foreach ($this->fullReportData as $userName => $dailyData) {
            // إنشاء كائن UserSheet لكل مستخدم
            // كل شيت يحمل اسم المستخدم وبياناته اليومية
            $sheets[] = new UserSheet($userName, $dailyData);
        }

        return $sheets;
    }
}

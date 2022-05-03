<?php

namespace Sezamin\IOExcel;
/**
 * Column
 */
class ExcelColumn
{
    public $label;
    public $index;
    public $isLinkedList = false;
    public $values = [];
    public $autoSize = true;
    public $size = 40;

    public function getColumnName(): string
    {
        return self::columnNumberToLetter($this->index);
    }

    public static function columnNumberToLetter(int $number): string
    {
        $numeric = ($number - 1) % 26;
        $letter = chr(65 + $numeric);
        $num = intval(($number - 1) / 26);
        if ($num > 0) {
            return self::columnNumberToLetter($num) . $letter;
        } else {
            return $letter;
        }
    }
}
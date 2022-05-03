<?php

namespace Sezamin\IOExcel;
use Matrix\Exception;
use phpDocumentor\Reflection\Types\Nullable;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;


class ExcelExport
{
    public $sheetName = 'Data';
    public $linkedSheetName = 'LinkedData';
    /** @var array */
    public $columns;
    /** @var array */
    public $columnsByGroup;
    /** @var array */
    public $groups;
    /** @var array */
    public $data;
    /** @var array */
    public $withGroup = false;
    /** @var boolean */
    public $ignoreEmptyRows = false;
    /** @var Worksheet */
    public $sheet = null;
    /** @var Worksheet */
    public $linkedData = null;
    public $groupFontSize = 16;

    function __construct($columns, $data = [], $withGroup = true, $ignoreEmptyRows = false){
        $this->withGroup = $withGroup;
        $this->ignoreEmptyRows = $ignoreEmptyRows;
        $this->prepareColumns($columns);
        $this->prepareData($data);
    }

    /** Export to xlsx file
     * @param $xlsxFileName
     * @return bool
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function export($xlsxFileName){
        try {
            $this->initSheets();
            $this->setHeaders();
            $this->setLinkedData();
            $this->setData();
            $writer = new Xlsx($this->xlsx);
            $writer->save($xlsxFileName);
            return true;
        }catch (Exception $e){
            throw new Exception("Error on create xlsx file");
        }

    }

    /** Preparing columns for export
     * @param $columns
     * @return void
     */

    private function prepareColumns($columns){
        $columnsByGroup = [];
        foreach($columns as $column){
            $params = is_string($column) ? ['label'=>$column]: $column;
            $labelGroup = $params['labelGroup'] ?? '';
            $columnsByGroup[$labelGroup][] = $params;
        }

        $colIndex = 1;
        foreach($columnsByGroup as $groupColumns){
            foreach($groupColumns as $column){
                $col = new ExcelColumn();
                $col->index = $colIndex;
                $col->label = $column['label'] ?? '';
                if(
                    isset($column['values'])
                    && is_array($column['values'])
                    && count($column['values'])>0
                ){
                    $col->values = $column['values'];
                    $col->isLinkedList = true;
                }
                if(isset($column['size'])){
                    $col->size = (int)$column['size'];
                }else{
                    $col->autoSize = true;
                }
                $this->columns[] = $col;
                $colIndex++;
            }
        }

        $this->columnsByGroup = $columnsByGroup;
    }

    /** Preparing for export
     * @param $data
     * @return void
     */
    private function prepareData($data)
    {
        foreach($data as $row){
            if(is_array($row) && count($row)>0){
                $this->data[] = $row;
            }else{
                if(!$this->ignoreEmptyRows){
                    $this->data[] = [];
                }
            }
        }
    }

    /** Init Xlsx sheets Data and Linked Data
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function initSheets(){
        $this->xlsx = new Spreadsheet();
        $this->xlsx->removeSheetByIndex(0);

        $this->sheet = new Worksheet(null, $this->sheetName);
        $this->xlsx->addSheet($this->sheet);

        $this->linkedData = new Worksheet(null, $this->linkedSheetName);
        $this->xlsx->addSheet($this->linkedData);
    }

    /** Set excel headers
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setHeaders(){
        $this->xlsx->setActiveSheetIndex(0);
        $rowIndex = 1;
        $sheet = $this->sheet;
        if($this->withGroup){
            $colIndex = 1;
            $sheet->getRowDimension($rowIndex)->setRowHeight(25);
            foreach($this->columnsByGroup as $labelGroup => $groupItems){
                $groupItemsCount = count($groupItems);
                $colName = ExcelColumn::columnNumberToLetter($colIndex) . "{$rowIndex}";
                $colToName = ExcelColumn::columnNumberToLetter($colIndex + $groupItemsCount - 1). "{$rowIndex}";
                $sheet->mergeCells("{$colName}:" . "{$colToName}");
                $sheet->setCellValue($colName, $labelGroup);
                $sheet->getStyle($colName)->getFont()->setSize($this->groupFontSize);
                $this->setHeaderCellFormat($sheet, "{$colName}:" . "{$colToName}");

                $colIndex += $groupItemsCount;
            }
            $rowIndex++;
        }

        $this->setHeaderRow($this->sheet, 2);
        $this->setHeaderRow($this->linkedData, 1);

    }

    private function setHeaderRow($sheet, $rowIndex){
        foreach($this->columns as $col){
            $colName = $col->getColumnName() . "{$rowIndex}";
            $sheet->setCellValue($colName, $col->label);
            if($col->autoSize){
                $sheet->getColumnDimensionByColumn($col->index)->setAutoSize(true);
            }
            $this->setHeaderCellFormat($sheet, $colName);
        }
    }

    private function setHeaderCellFormat($sheet, $colName){
        $sheet->getStyle($colName)->getFont()->setBold(true);
        $sheet->getStyle($colName)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($colName)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle($colName)->getBorders()->getOutline()
            ->setBorderStyle(Border::BORDER_HAIR)
            ->setColor(new Color('00000000'));
    }

    /** Set values
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setData(){
        $this->xlsx->setActiveSheetIndex(0);
        $sheet = $this->sheet;
        $rowIndex = $this->withGroup ? 2 : 1;
        foreach($this->data as $row){
            $rowIndex++;
            if(count($row)>0){
                foreach ($this->columns as $col){
                    $colName = $col->getColumnName() . "{$rowIndex}";
                    $cellValue = $row[$col->index - 1] ?? '';
                    $sheet->setCellValue($colName, $cellValue);
                    if($col->autoSize){
                        $sheet->getColumnDimensionByColumn($col->index)->setAutoSize(true);
                    }
                    $sheet->getStyle($colName)->getBorders()->getOutline()
                        ->setBorderStyle(Border::BORDER_HAIR)
                        ->setColor(new Color('00000000'));
                }
            }
        }
    }

    /** Set all linked data and link values to cells
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setLinkedData(){
        $linkedData = $this->linkedData;
        $sheet = $this->sheet;
        foreach($this->columns as $col){
            $this->xlsx->setActiveSheetIndex(1);
            $rowIndex = 2;
            $colName = $col->getColumnName();
            if(is_array($col->values) && count($col->values)>0){
                foreach($col->values as $value){
                    $linkedData->setCellValue($colName. "{$rowIndex}", $value);
                    $rowIndex++;
                }
                if($col->autoSize){
                    $linkedData->getColumnDimensionByColumn($col->index)->setAutoSize(true);
                }
                $this->xlsx->setActiveSheetIndex(0);
                $this->setDataValidationList($sheet, $colName, 1000, $rowIndex);
            }
        }
    }

    /** Set validation list for cell
     * @param $sheet
     * @param $colName
     * @param $toRowIndex
     * @param $linkToRowIndex
     * @param $promptText
     * @param $promptTitle
     * @return void
     */
    public function setDataValidationList($sheet, $colName, $toRowIndex, $linkToRowIndex, $promptText = '', $promptTitle = '')
    {
        $fromIndex = $this->withGroup ? 3 :2;

        $validation = $sheet->getCell("{$colName}{$fromIndex}")->getDataValidation();
        $validation->setType( DataValidation::TYPE_LIST );
        $validation->setErrorStyle( DataValidation::STYLE_INFORMATION );
        $validation->setAllowBlank(false);
        $validation->setShowInputMessage(true);
        $validation->setShowDropDown(true);
        if(!empty($promptTitle)){
            $validation->setPromptTitle($promptTitle);
            if(!empty($promptText)){
                $validation->setPrompt($promptText);
            }
        }

        $validation->setFormula1("{$this->linkedSheetName}!\${$colName}\$2:\${$colName}\${$linkToRowIndex}");

        $sheet->setDataValidation("{$colName}{$fromIndex}:{$colName}{$toRowIndex}", $validation);
    }
}
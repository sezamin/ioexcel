<?php

use PHPUnit\Framework\TestCase;
use Sezamin\IOExcel\ExcelExport;

final class ExportExcelTest extends TestCase
{
    public function testExport(): void
    {
        $testTmp = __DIR__ . "/results/testArray.xlsx";

        $e = new ExcelExport(
            ["N", "Name", "Rating", "Employers", "Customers"],
            $data = [
                [1, "Company Name 1", 1000, 10, 1020],
                [2, "Company Name 2", 1001, 20, 1010],
                [3, "Company Name 3", 2000, 23, 3000],
                [4, "Company Name 4", 2100, 12, 302],
                [5, "Company Name 5", 1034, 65, 1245],
            ],
            false);

        $e->export($testTmp);
        $this->assertEquals($data, $data, 'Error on array export');
    }


    public function testExportGroup(): void
    {
        $testTmp = __DIR__ . "/results/testGroup.xlsx";
        $data = [];
        $faker = Faker\Factory::create();
        for($i=1; $i< 100; $i++){
            $data[] = [$i, $faker->name(), $faker->randomNumber(5), $faker->randomNumber(5), $faker->randomNumber(5)];
        }
        $e = new ExcelExport(
            [
                ['label'=>"N", 'labelGroup'=>"Base"],
                ['label'=>"Name", 'labelGroup'=>"Base"],
                ['label'=>"Rating", 'labelGroup'=>"Rate"],
                ['label'=>"Employers", 'labelGroup'=>"Rate"],
                ['label'=>"Customers", 'labelGroup'=>"Rate"]
            ],
            $data);

        $e->export($testTmp);

        $this->assertEquals($data, $data, 'Error on array export with group');


    }

    public function testExportWithLinkedData(): void
    {
        $testTmp = __DIR__ . "/results/testLinked.xlsx";
        $data = [];
        $faker = Faker\Factory::create();
        $linked1Count = 10;
        $linked2Count = 30;
        $linked1Data = [];
        $linked2Data = [];
        for($i=1; $i<= $linked1Count; $i++){
            $linked1Data[] = $faker->company();
        }
        for($i=1; $i<= $linked2Count; $i++){
            $linked2Data[] = $faker->domainName();
        }

        for($i=1; $i< 200; $i++){
            $data[] = [
                $i, $faker->name(), $faker->randomNumber(5), $faker->randomNumber(5), $faker->randomNumber(5),
                $linked1Data[rand(1, $linked1Count) -1],
                $linked2Data[rand(1, $linked2Count) -1],
            ];
        }

        $e = new ExcelExport(
            [
                ['label'=>"Num", 'labelGroup'=>"Base"],
                ['label'=>"Name", 'labelGroup'=>"Base"],
                ['label'=>"Rating", 'labelGroup'=>"Rate"],
                ['label'=>"Employers", 'labelGroup'=>"Rate"],
                ['label'=>"Customers", 'labelGroup'=>"Rate"],
                ['label'=>"Linked Data 1", 'labelGroup'=>"Linked", "values"=>$linked1Data],
                ['label'=>"Linked Data 2", 'labelGroup'=>"Linked", "values"=>$linked2Data],
            ],
            $data);

        $e->export($testTmp);

        $this->assertEquals($data, $data, 'Error on array export with group');

    }

}
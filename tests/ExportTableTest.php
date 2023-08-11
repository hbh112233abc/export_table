<?php
namespace Bingher\ExportTableTest;

use Bingher\ExportTable\ExportTable;
use PHPUnit\Framework\TestCase;

class ExportTableTest extends TestCase
{
    public function testColumnLabel(): void
    {
        $a = ExportTable::columnLabel(0);
        $this->assertEquals($a, 'A');
        $g = ExportTable::columnLabel(6);
        $this->assertEquals($g, 'G');
        $aa = ExportTable::columnLabel(26);
        $this->assertEquals($aa, 'AA');
        $ab = ExportTable::columnLabel(27);
        $this->assertEquals($ab, 'AB');
    }

    public function testIsImage(): void
    {
        $n = ExportTable::isImage('');
        $this->assertEquals($n, false);
        $file = tempnam('.', uniqid());
        $n    = ExportTable::isImage($file);
        $this->assertEquals($n, false);

    }
}

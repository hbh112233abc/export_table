<?php
namespace Bingher\ExportTable;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExportTable
{
    protected $excel;
    protected $activeSheet;

    protected $headers;

    protected $rowHeight = 80;

    public function __construct(int $rowHeight = 80)
    {
        $this->excel       = new Spreadsheet();
        $this->activeSheet = $this->excel->getActiveSheet();
        $this->rowHeight   = $rowHeight;
    }

    public function header(string $column, string $name, array $style = [])
    {
        $defaultStyle = [
            'vertical'   => Alignment::VERTICAL_CENTER,
            'horizontal' => Alignment::HORIZONTAL_GENERAL,
        ];
        $style        = array_merge($defaultStyle, $style);
        if (empty($this->headers[$column])) {
            $this->headers[$column] = [
                'name'  => $name,
                'style' => $style,
            ];
        } else {
            $this->headers[$column]['style'] = array_merge(
                $this->headers[$column]['style'],
                $style
            );
        }
    }

    public function data(array $data, array $headerNames = [])
    {
        $keys    = array_keys($data[0]);
        $columns = [];
        foreach ($keys as $index => $key) {
            $columnName    = empty($headerNames[$index]) ? $key : $headerNames[$index];
            $columnLabel   = $this->columnLabel($index);
            $columns[$key] = $columnLabel;
            $this->header($columnLabel, $columnName);
        }

        # 生成头部
        $this->createHeader();

        # 加入数据
        foreach ($data as $index => $rowData) {
            $row = $index + 2;
            foreach ($rowData as $key => $value) {
                $col  = $columns[$key];
                $cell = $col . $row;

                if ($this->isImage($value)) {
                    $this->insertImage($cell, $value);
                } else {
                    $this->activeSheet->setCellValue($cell, $value);
                }
            }
            $this->activeSheet->getRowDimension($row)->setRowHeight($this->rowHeight);
        }

        return $this;
    }

    public static function isImage($file)
    {
        if (!file_exists($file)) {
            return false;
        }
        $allowedFileExtensions = array('jpg', 'jpeg', 'png', 'gif', /* 添加其他图片文件扩展名 */);

        $fileExtension = pathinfo($file, PATHINFO_EXTENSION);

        return in_array(strtolower($fileExtension), $allowedFileExtensions);
    }

    public function createHeader()
    {
        foreach ($this->headers as $column => $header) {
            $style               = $header['style'] ?? [];
            $alignmentHorizontal = $style['horizontal'] ?? Alignment::HORIZONTAL_GENERAL;
            $this->activeSheet->getStyle($column)->getAlignment()->setHorizontal($alignmentHorizontal);
            $alignmentVertical = $style['vertical'] ?? Alignment::VERTICAL_CENTER;
            $this->activeSheet->getStyle($column)->getAlignment()->setVertical($alignmentVertical);
            if (!empty($style['width'])) {
                $this->activeSheet->getColumnDimension($column)->setWidth($style['width']);
            }
            $this->activeSheet->setCellValue($column . '1', $header['name']);
        }
    }

    public function insertImage(
        string $cell,
        string $image,
        array $size = [80, 80],
        array $offset = [12, 12]
    ) {
        if (!file_exists($image)) {
            return;
        }
        $img = new Drawing();
        $img->setPath($image);
        // 设置宽度高度
        $img->setHeight($size[0]); //照片高度
        $img->setWidth($size[1]); //照片宽度
        /*设置图片要插入的单元格*/
        $img->setCoordinates($cell);
        // 图片偏移距离
        $img->setOffsetX($offset[0]);
        $img->setOffsetY($offset[1]);
        $img->setWorksheet($this->activeSheet);
    }

    public static function columnLabel(int $index)
    {
        $chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        if ($index < 26) {
            return $chars[$index];
        } else {
            $x = intval(floor($index / 26));
            $m = $index % 26;
            return $chars[$x - 1] . $chars[$m];
        }
    }

    public function output(string $filename = "")
    {
        ob_clean();
        $filename = iconv("utf-8", "gb2312", $filename);
        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition: attachment;filename=\"$filename\"");
        header('Cache-Control: max-age=0');
        $writer = new Xlsx($this->excel);
        $writer->save('php://output'); //文件通过浏览器下载
    }
}

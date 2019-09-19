<?php
/**
 * Created by PhpStorm.
 * User: wen
 * Date: 2019/9/18
 * Time: 16:15
 * excel
 */

namespace App\Admin\Services\phpExcel;


class WenExcel
{
    static protected $phpExcelObj = null; // phpExcel对象

    static protected $creator = 'wen-create'; // 文档创建者

    static protected $docTitle = 'wen-title'; // 文档标题

    static protected $modifier = 'wen-modify'; // 文档修改者

    static protected $subject = 'wen-subject'; // 文档科目

    static protected $description = 'wen-description'; // 文档描述

    static protected $keyword = 'wen-keyword'; // 关键词

    static protected $gategory = 'wen-gategory'; // 文档类别

    static protected $textAlign = 'center'; // 水平对齐

    static protected $verticalAlign = 'center'; // 垂直对齐

    static protected $titleHeight = 30; // 标题高度

    static protected $contentHeight = 25; // 内容高度

    static protected $cellWidth = []; // 单元格宽度设置

    static protected $title = ''; // sheet标题

    static protected $filename = ''; // 文件名称

    static protected $header = []; // 头信息集合

    static protected $data = []; // 数据集合

    static protected $line = 1; // 起始行数

    static protected $page = 0; // 起始页数

    static protected $perPageNum = 0; // 每页条数，如果大于0，则分页，否则不分页

    static protected $formatCallback = null; // 数据格式化回调

    static protected $body = [];

    static protected $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;

    static protected $cacheSettings = [];

    // 列
    static protected $letter = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T', 'U','V','W','X','Y','Z',
        'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ'];

    function __construct()
    {
        self::init();
    }


    public static function saveFile($type = 'xls')
    {
        $filename = __DIR__.'/'.self::$filename.'.'.$type;
//        dd($filename);
        self::setData();
        $objWriter = \PHPExcel_IOFactory::createWriter(self::getPhpExcelObj(), 'Excel5');
        $objWriter->save($filename);
    }

    /**
     * @param string $type
     * @throws \PHPExcel_Reader_Exception
     * @throws \PHPExcel_Writer_Exception
     * 执行导出
     */
    public static function export($type = 'xls')
    {
        set_time_limit(0);
        header('Content-Type: application/vnd.ms-excel');
//        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.self::$filename.'.'.$type.'"');
        header('Cache-Control: max-age=0');
//        application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
        \PHPExcel_Settings::setCacheStorageMethod(self::$cacheMethod, self::$cacheSettings);
        self::setData();
        $objWriter = \PHPExcel_IOFactory::createWriter(self::getPhpExcelObj(), 'Excel5');
        $objWriter->save('php://output');
    }

    /**
     * @return int
     * 获取当前工作表的当前行
     */
    public static function getLine()
    {
        return self::$line;
    }

    public static function getLetter()
    {
        return self::$letter;
    }

    /**
     * @param $row
     * @param null $col_start
     * @param null $col_end
     * 合并单元格
     */
    public static function mergeCells($row, $col_start = null, $col_end = null)
    {
        if (empty($col_start)) {
            $col_start = self::getFirstLetter();
        }
        if (empty($col_end)) {
            $col_end = self::getLastLetter();
        }
        self::getPhpExcelObj()->getActiveSheet()->mergeCells($col_start.$row.':'.$col_end.$row);
    }

    /**
     * @param $header
     * @param $data
     * @param $filename
     * @param int $pageNum
     * @param string $sheetTitle
     * 设置头信息集合、数据集合、文件名称、sheet标题和每页数据条数
     */
    public static function setProperties($header, $data,  $filename, $sheetTitle = '', $pageNum = 0)
    {
        self::$header = $header;
        self::$data = $data;
        self::$filename = $filename.date('YmdHis');
        self::$perPageNum = $pageNum;
        self::$title = $sheetTitle ? $sheetTitle : $filename;
    }


    /**
     * 创建工作表
     */
    public static function addSheet()
    {
        self::setActiveSheetIndex(self::$page);
        self::$page++;
    }

    public static function setFormatCallback($func, $body)
    {
        self::$formatCallback = $func;
        self::$body = $body;
    }

    /**
     * @param $row
     * @param $col
     * @param $val
     * 通过外部设置数据集
     */
    public static function setDataByUser($row, $col, $val)
    {
        self::setCellValue($row, $col, $val);
        if (isset(self::$cellWidth[$col])) {
            self::setWidth($col, self::$cellWidth[$col]);
        }
        self::setTextAlign($col, $row);
        self::setVerticalAlign($col, $row);
        self::setHeight($row, self::$contentHeight);
    }

    /**
     * 设置数据体
     */
    protected static function setData()
    {
        if (self::$perPageNum > 0) {
            self::$data = array_chunk(self::$data, self::$perPageNum);
            foreach (self::$data as $key => $item) {
                self::makeSheet($item);
                self::createSheet();
            }
        } else {
            self::makeSheet(self::$data);
        }

    }


    /**
     * @return null
     * 获取PHPExcel对象
     */
    protected static function getPhpExcelObj()
    {
        if (self::$phpExcelObj == null) {
            self::init();
            self::setDocProperties();
        }
        return self::$phpExcelObj;
    }


    /**
     * @param $data
     * 工作表数据填写
     */
    protected static function makeSheet($data)
    {
        self::addSheet();
        self::setHeader();
        $row = self::$line;
        foreach ($data as $key => $item) {
            if (self::$formatCallback) {
                $item = call_user_func(self::$formatCallback, $item, self::$body);
            }
            foreach ($item as $k => $v) {
                $col = self::$letter[$k];
                self::setCellValue($row, $col, $v);
                if (isset(self::$cellWidth[$col])) {
                    self::setWidth($col, self::$cellWidth[$col]);
                }
                self::setTextAlign($col, $row);
                self::setVerticalAlign($col, $row);
            }
            self::setHeight($row, self::$contentHeight);
            $row++;
        }
        self::setSheetTitle(self::$title.self::$page);
    }

    /**
     * 创建一个sheet
     */
    protected static function createSheet()
    {
        self::getPhpExcelObj()->createSheet();
    }

    /**
     * @param $title
     * 设置sheet标题
     */
    protected static function setSheetTitle($title)
    {
        self::getPhpExcelObj()->getActiveSheet()->setTitle($title);
    }

    /**
     * 设置文件头
     */
    protected static function setHeader()
    {
        foreach (self::$header as $key => $item) {
            $col = self::$letter[$key];
            self::setCellValue(self::$line, $col, $item);
            if (isset(self::$cellWidth[$col])) {
                self::setWidth($col, self::$cellWidth[$col]);
            }
            self::setTextAlign($col, self::$line);
            self::setVerticalAlign($col, self::$line);
            self::setFontBold($col, self::$line);
        }
        self::setHeight(self::$line, self::$titleHeight);
        self::$line++;
    }


    /**
     * @param int $index
     * 设置工作表
     */
    protected static function setActiveSheetIndex($index = 0) {
        self::getPhpExcelObj()->setActiveSheetIndex($index);
    }

    /**
     * @param $row
     * @param $col
     * @param $value
     * 设置单元格的值
     */
    protected static function setCellValue($row, $col, $value)
    {
        self::getPhpExcelObj()->getActiveSheet()->setCellValue($col.$row, $value);
    }

    /**
     * @param $col
     * @param $row
     * 设置单元格粗体
     */
    protected static function setFontBold($col, $row)
    {
        self::getPhpExcelObj()->getActiveSheet()->getStyle($col.$row)->getFont()->setBold(true);
    }

    /**
     * @param $col
     * @param $row
     * @param $size
     * 设置单元格字体大小
     */
    protected static function setFontSize($col, $row, $size)
    {
        self::getPhpExcelObj()->getActiveSheet()->getStyle($col.$row)->getFont()->setSize($size);
    }

    /**
     * @param $col
     * @param $row
     * 设置单元格水平对其方式
     */
    protected static function setTextAlign($col, $row)
    {
        self::getPhpExcelObj()->getActiveSheet()->getStyle($col.$row)->getAlignment()->setHorizontal(self::getTextAlign());
    }

    /**
     * @param $col
     * @param $row
     * 设置单元格垂直对齐方式
     */
    protected static function setVerticalAlign($col, $row)
    {
        self::getPhpExcelObj()->getActiveSheet()->getStyle($col.$row)->getAlignment()->setVertical(self::getVerticalAlign());
    }

    /**
     * @param $col
     * @param $width
     * 设置单元格宽度
     */
    protected static function setWidth($col, $width)
    {
        self::getPhpExcelObj()->getActiveSheet()->getColumnDimension($col)->setWidth($width);
    }

    /**
     * @param $row
     * @param $height
     * 设置单元格高度
     */
    protected static function setHeight($row, $height)
    {
        self::getPhpExcelObj()->getActiveSheet()->getRowDimension($row)->setRowHeight($height);
    }

    /**
     * 初始化配置信息
     */
    public static function init()
    {
        $configs = require_once __DIR__.DIRECTORY_SEPARATOR.'config.php';
        self::$creator = isset($configs['creator']) ? $configs['creator'] : 'wen-creator';
        self::$modifier = isset($configs['modifier']) ? $configs['modifier'] : 'wen-modifier';
        self::$subject = isset($configs['subject']) ? $configs['subject'] : 'wen-subject';
        self::$description = isset($configs['description']) ? $configs['description'] : 'wen-description';
        self::$keyword = isset($configs['keyword']) ? $configs['keyword'] : 'wen-keyword';
        self::$docTitle = isset($configs['docTitle']) ? $configs['docTitle'] : 'wen-docTitle';
        self::$gategory = isset($configs['gategory']) ? $configs['gategory'] : 'wen-gategory';
        self::$titleHeight = isset($configs['titleHeight']) ? $configs['titleHeight'] : 30;
        self::$contentHeight = isset($configs['contentHeight']) ? $configs['contentHeight'] : 25;
        self::$cellWidth = isset($configs['cellWidth']) ? $configs['cellWidth'] : 25;
        self::$page = 0;
        self::$line = 1;
        self::$phpExcelObj = new \PHPExcel();
    }

    /**
     * @return mixed
     * 获取第一列
     */
    protected static function getFirstLetter()
    {
        return self::$letter[0];
    }

    /**
     * @return mixed
     * 获取最后一列
     */
    protected static function getLastLetter()
    {
        if (count(self::$header) > 0) {
            return self::$letter[count(self::$header) - 1];
        }
        return 0;
    }

    /**
     * 设置文档属性
     */
    protected static function setDocProperties()
    {
        self::$phpExcelObj->getProperties()->setCreator(self::$creator)
            ->setLastModifiedBy(self::$modifier)
            ->setTitle(self::$docTitle)
            ->setSubject(self::$subject)
            ->setDescription(self::$description)
            ->setKeywords(self::$keyword)
            ->setCategory(self::$gategory);
    }

    /**
     * @return string
     * 获取水平对齐方式
     */
    protected static function getTextAlign()
    {
        switch (self::$textAlign) {
            case 'center':
                return \PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
                break;
            case 'right':
                return \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
                break;
            default:
                return \PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
        }
    }

    /**
     * @return string
     * 获取垂直对齐方式
     */
    protected static function getVerticalAlign()
    {
        switch (self::$verticalAlign) {
            case 'center':
                return \PHPExcel_Style_Alignment::VERTICAL_CENTER;
                break;
            case 'top':
                return \PHPExcel_Style_Alignment::VERTICAL_TOP;
                break;
            default:
                return \PHPExcel_Style_Alignment::VERTICAL_BOTTOM;
        }
    }

}
<?php
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

App::uses('AppHelper', 'View/Helper');

/**
 * Helper for working with PHPExcel class.
 *
 * @package PhpExcel
 * @author segy
 */
class PhpExcelHelper extends AppHelper {
    /**
     * Instance of PHPExcel class
     *
     * @var PHPExcel
     */
    protected $_xls;

    /**
     * Pointer to current row
     *
     * @var int
     */
    protected $_row = 1;

    /**
     * Internal table params
     *
     * @var array
     */
    protected $_tableParams;

    /**
     * Number of rows
     *
     * @var int
     */
    protected $_maxRow = 0;

    /**
     * Create new worksheet or load it from existing file
     *
     * @return $this for method chaining
     */
    public function createWorksheet() {
        $this->_xls = new Spreadsheet();
        $this->_row = 1;

        return $this;
    }

    /**
     * Create new worksheet from existing file
     *
     * @param string $file path to excel file to load
     * @return $this for method chaining
     */
    public function loadWorksheet($file) {
        $this->_xls = PHPExcel_IOFactory::load($file);
        $this->setActiveSheet(0);
        $this->_row = 1;

        return $this;
    }

    /**
     * Add sheet
     *
     * @param string $name
     * @return $this for method chaining
     */
    public function addSheet($name) {
        $index = $this->_xls->getSheetCount();
        $this->_xls->createSheet($index)
            ->setTitle($name);

        $this->setActiveSheet($index);

        return $this;
    }

    /**
     * Set active sheet
     *
     * @param int $sheet
     * @return $this for method chaining
     */
    public function setActiveSheet($sheet) {
        $this->_maxRow = $this->_xls->setActiveSheetIndex($sheet)->getHighestRow();
        $this->_row = 1;

        return $this;
    }

    /**
     * Set worksheet name
     *
     * @param string $name name
     * @return $this for method chaining
     */
    public function setSheetName($name) {
        $this->_xls->getActiveSheet()->setTitle($name);

        return $this;
    }

    /**
     * Overloaded __call
     * Move call to PHPExcel instance
     *
     * @param string function name
     * @param array arguments
     * @return the return value of the call
     */
    public function __call($name, $arguments) {
        return call_user_func_array(array($this->_xls, $name), $arguments);
    }

    /**
     * Set default font
     *
     * @param string $name font name
     * @param int $size font size
     * @return $this for method chaining
     */
    public function setDefaultFont($name, $size) {
        $this->_xls->getParent()->getDefaultStyle()->getFont()->setName($name);
        $this->_xls->getParent()->getDefaultStyle()->getFont()->setSize($size);

        return $this;
    }

    /**
     * Set row pointer
     *
     * @param int $row number of row
     * @return $this for method chaining
     */
    public function setRow($row) {
        $this->_row = (int)$row;

        return $this;
    }

    /**
     * Start table - insert table header and set table params
     *
     * @param array $data data with format:
     *   label   -   table heading
     *   width   -   numeric (leave empty for "auto" width)
     *   filter  -   true to set excel filter for column
     *   wrap    -   true to wrap text in column
     * @param array $params table parameters with format:
     *   offset  -   column offset (numeric or text)
     *   font    -   font name of the header text
     *   size    -   font size of the header text
     *   bold    -   true for bold header text
     *   italic  -   true for italic header text
     * @return $this for method chaining
     */
    public function addTableHeader($data, $params = array()) {
        // offset
        $offset = 1;
        if (isset($params['offset']))
            $offset = is_numeric($params['offset']) ? (int)$params['offset'] : Coordinate::columnIndexFromString($params['offset']);

        // font name
        if (isset($params['font']))
            $this->_xls->getActiveSheet()->getStyle($this->_row)->getFont()->setName($params['font']);

        // font size
        if (isset($params['size']))
            $this->_xls->getActiveSheet()->getStyle($this->_row)->getFont()->setSize($params['size']);

        // bold
        if (isset($params['bold']))
            $this->_xls->getActiveSheet()->getStyle($this->_row)->getFont()->setBold($params['bold']);

        // italic
        if (isset($params['italic']))
            $this->_xls->getActiveSheet()->getStyle($this->_row)->getFont()->setItalic($params['italic']);

        if (isset($params['color']))
            $this->_xls->getActiveSheet()->getStyle($this->_row)->getFont()->applyFromArray(array(
                'color' => array(
                    'rgb' => $params['color'])
            ));

        if (isset($params['fill']))
            $this->_xls->getActiveSheet()->getStyle($this->_row)->applyFromArray(array(
                'fill' => array(
                    'fillType' => Fill::FILL_SOLID,
                    'color' => array('rgb' => $params['fill']))
            ));

        if (isset($params['formatting'])) {
            foreach ($params['formatting'] as $col => $value) {
                $this->_xls->getActiveSheet()->getStyle($col)->getNumberFormat()->setFormatCode($value);
            }
        }

        // set internal params that need to be processed after data are inserted
        $this->_tableParams = array(
            'header_row' => $this->_row,
            'offset' => $offset,
            'row_count' => 0,
            'auto_width' => array(),
            'filter' => array(),
            'wrap' => array(),
            'email' => array()
        );

        foreach ($data as $d) {
            // set label
            $this->_xls->getActiveSheet()->setCellValueByColumnAndRow($offset, $this->_row, $d['label']);

            // set width
            if (isset($d['width']) && is_numeric($d['width']))
                $this->_xls->getActiveSheet()->getColumnDimensionByColumn($offset)->setWidth((float)$d['width']);
            else
                $this->_tableParams['auto_width'][] = $offset;

            // filter
            if (isset($d['filter']) && $d['filter'])
                $this->_tableParams['filter'][] = $offset;

            // wrap
            if (isset($d['wrap']) && $d['wrap'])
                $this->_tableParams['wrap'][] = $offset;

            // email
            if (isset($d['email']) && $d['email'])
                $this->_tableParams['email'][] = $offset;

            $offset++;
        }
        $this->_row++;

        return $this;
    }

    /**
     * Write array of data to current row
     *
     * @param array $data
     * @return $this for method chaining
     */
    public function addTableRow($data) {
        $offset = $this->_tableParams['offset'];

        foreach ($data as $d)
            $this->_xls->getActiveSheet()->setCellValueByColumnAndRow($offset++, $this->_row, $d);

        $this->_row++;
        $this->_tableParams['row_count']++;

        return $this;
    }

    /**
     * End table - set params and styles that required data to be inserted first
     *
     * @return $this for method chaining
     */
    public function addTableFooter() {
        // auto width
        foreach ($this->_tableParams['auto_width'] as $col)
            $this->_xls->getActiveSheet()->getColumnDimensionByColumn($col)->setAutoSize(true);

        // filter (has to be set for whole range)
        if (count($this->_tableParams['filter']))
            $this->_xls->getActiveSheet()->setAutoFilter(Coordinate::stringFromColumnIndex($this->_tableParams['filter'][0]) . ($this->_tableParams['header_row']) . ':' . Coordinate::stringFromColumnIndex($this->_tableParams['filter'][count($this->_tableParams['filter']) - 1]) . ($this->_tableParams['header_row'] + $this->_tableParams['row_count']));

        // wrap
        foreach ($this->_tableParams['wrap'] as $col)
            $this->_xls->getActiveSheet()->getStyle(Coordinate::stringFromColumnIndex($col) . ($this->_tableParams['header_row'] + 1) . ':' . Coordinate::stringFromColumnIndex($col) . ($this->_tableParams['header_row'] + $this->_tableParams['row_count']))->getAlignment()->setWrapText(true);

        //jwallaced added email links and formatting
        foreach ($this->_tableParams['email'] as $col) {
            $lastRow = $this->_xls->getActiveSheet()->getHighestRow();
            for ($row = 2; $row <= $lastRow; $row++) {
                $cell = $this->_xls->getActiveSheet()->getCell(Coordinate::stringFromColumnIndex($col) . $row);
                //add formatting
                $this->_xls->getActiveSheet()->getStyle(Coordinate::stringFromColumnIndex($col) . $row)->applyFromArray(array(
                    'font'  => array(
                        'color' => array('rgb' => '0000FF'),
                        'underline' => true)
                ));
                //add hyperlinks
                $cell->getHyperlink($cell->getValue())->setUrl('mailto:'. $cell->getValue() );
            }
        }

        return $this;
    }

    /**
     * Write array of data to current row starting from column defined by offset
     *
     * @param array $data
     * @return $this for method chaining
     */
    public function addData($data, $offset = 1) {
        // solve textual representation
        if (!is_numeric($offset))
            $offset = Coordinate::columnIndexFromString($offset);

        foreach ($data as $d)
            $this->_xls->getActiveSheet()->setCellValueByColumnAndRow($offset++, $this->_row, $d);

        $this->_row++;

        return $this;
    }

    /**
     * Get array of data from current row
     *
     * @param int $max
     * @return array row contents
     */
    public function getTableData($max = 100) {
        if ($this->_row > $this->_maxRow)
            return false;

        $data = array();

        for ($col = 0; $col < $max; $col++)
            $data[] = $this->_xls->getActiveSheet()->getCellByColumnAndRow($col, $this->_row)->getValue();

        $this->_row++;

        return $data;
    }

    /**
     * Get writer
     *
     * @return \PhpOffice\PhpSpreadsheet\Writer\IWriter
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function getWriter() {
        return IOFactory::createWriter($this->_xls, 'Xlsx');
    }

    /**
     * Save to a file
     *
     * @param string $file path to file
     * @return bool
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function save($file) {
        $objWriter = $this->getWriter();
        return $objWriter->save($file);
    }

    /**
     * Output file to browser
     *
     * @param string $filename
     * @return void on this call
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function output($filename = 'export.xlsx') {
        // remove all output
        ob_end_clean();

        // headers
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
        header('Cache-Control: max-age=0');

        // writer
        $objWriter = $this->getWriter();
        $objWriter->save('php://output');

        exit;
    }

    /**
     * Free memory
     *
     * @return void
     */
    public function freeMemory() {
        $this->_xls->disconnectWorksheets();
        unset($this->_xls);
    }
}

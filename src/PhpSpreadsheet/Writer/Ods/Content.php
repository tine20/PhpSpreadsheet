<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Ods;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Shared\XMLWriter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Worksheet\ColumnDimension;
use PhpOffice\PhpSpreadsheet\Worksheet\Row;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Ods\Cell\Comment;

/**
 * @category PhpSpreadsheet
 *
 * @method Ods getParentWriter
 *
 * @copyright Copyright (c) 2006 - 2015 PhpSpreadsheet (https://github.com/PHPOffice/PhpSpreadsheet)
 * @author    Alexander Pervakov <frost-nzcr4@jagmort.com>
 */
class Content extends WriterPart
{
    const NUMBER_COLS_REPEATED_MAX = 1024;
    const NUMBER_ROWS_REPEATED_MAX = 1048576;
    const CELL_STYLE_PREFIX = 'ce';
    const COLUMN_STYLE_PREFIX = 'co';
    const ROW_STYLE_PREFIX = 'ro';
    const TABLE_STYLE_PREFIX = 'ta';

    /**
     * Write content.xml to XML format.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     *
     * @return string XML Output
     */
    public function write()
    {
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8');

        // Content
        $objWriter->startElement('office:document-content');
        $objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
        $objWriter->writeAttribute('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0');
        $objWriter->writeAttribute('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
        $objWriter->writeAttribute('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
        $objWriter->writeAttribute('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
        $objWriter->writeAttribute('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0');
        $objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
        $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $objWriter->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
        $objWriter->writeAttribute('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0');
        $objWriter->writeAttribute('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
        $objWriter->writeAttribute('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0');
        $objWriter->writeAttribute('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0');
        $objWriter->writeAttribute('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0');
        $objWriter->writeAttribute('xmlns:math', 'http://www.w3.org/1998/Math/MathML');
        $objWriter->writeAttribute('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0');
        $objWriter->writeAttribute('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0');
        $objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
        $objWriter->writeAttribute('xmlns:ooow', 'http://openoffice.org/2004/writer');
        $objWriter->writeAttribute('xmlns:oooc', 'http://openoffice.org/2004/calc');
        $objWriter->writeAttribute('xmlns:dom', 'http://www.w3.org/2001/xml-events');
        $objWriter->writeAttribute('xmlns:xforms', 'http://www.w3.org/2002/xforms');
        $objWriter->writeAttribute('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema');
        $objWriter->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
        $objWriter->writeAttribute('xmlns:rpt', 'http://openoffice.org/2005/report');
        $objWriter->writeAttribute('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2');
        $objWriter->writeAttribute('xmlns:xhtml', 'http://www.w3.org/1999/xhtml');
        $objWriter->writeAttribute('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#');
        $objWriter->writeAttribute('xmlns:tableooo', 'http://openoffice.org/2009/table');
        $objWriter->writeAttribute('xmlns:field', 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0');
        $objWriter->writeAttribute('xmlns:formx', 'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0');
        $objWriter->writeAttribute('xmlns:css3t', 'http://www.w3.org/TR/css3-text/');
        $objWriter->writeAttribute('xmlns:loext', 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0');
        $objWriter->writeAttribute('office:version', '1.2');

        $objWriter->writeElement('office:scripts');
        $objWriter->writeElement('office:font-face-decls');

        // Styles XF
        $objWriter->startElement('office:automatic-styles');
        $this->writeXfStyles($objWriter, $this->getParentWriter()->getSpreadsheet());
        $objWriter->endElement();

        $objWriter->startElement('office:body');
        $objWriter->startElement('office:spreadsheet');
        $objWriter->writeElement('table:calculation-settings');

        $this->writeSheets($objWriter);

        $objWriter->writeElement('table:named-expressions');
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * Write sheets.
     *
     * @param XMLWriter $objWriter
     */
    private function writeSheets(XMLWriter $objWriter)
    {
        $spreadsheet = $this->getParentWriter()->getSpreadsheet(); // @var $spreadsheet Spreadsheet

        $sheetCount = $spreadsheet->getSheetCount();
        for ($i = 0; $i < $sheetCount; ++$i) {
            $sheet = $spreadsheet->getSheet($i);
            $objWriter->startElement('table:table');
            $objWriter->writeAttribute('table:name', $sheet->getTitle());
            if (null !== $sheet->getXfIndex() && isset($spreadsheet->getTableXfCollection()[$sheet->getXfIndex()])) {
                $objWriter->writeAttribute('table:style-name', self::TABLE_STYLE_PREFIX . $sheet->getXfIndex());
            }
            $objWriter->writeElement('office:forms');
            if (empty($sheet->getColumnDimensionCollection())) {
                $objWriter->startElement('table:table-column');
                $objWriter->writeAttribute('table:number-columns-repeated', self::NUMBER_COLS_REPEATED_MAX);
                $objWriter->endElement();
            } else {
                $this->writeColumns($objWriter, $sheet);
            }
            $this->writeRows($objWriter, $sheet);
            $objWriter->endElement();
        }
    }

    /**
     * Write columns of the specified sheet.
     *
     * @param XMLWriter $objWriter
     * @param Worksheet $sheet
     */
    private function writeColumns(XMLWriter $objWriter, Worksheet $sheet)
    {
        // @var ColumnDimension
        foreach ($sheet->getColumnDimensionCollection() as $columnDimension) {
            $objWriter->startElement('table:table-column');
            if ($columnDimension->getWidth() > 1) {
                $objWriter->writeAttribute('table:number-columns-repeated', $columnDimension->getWidth());
            }

            if (!empty($columnDimension->getXfIndex())) {
                $objWriter->writeAttribute(
                    'table:style-name',
                    self::COLUMN_STYLE_PREFIX .
                    $columnDimension->getXfIndex()
                );
            }

            if (!empty($columnDimension->getDefaultXfIndex())) {
                $objWriter->writeAttribute(
                    'table:default-cell-style-name',
                    self::CELL_STYLE_PREFIX .
                    $columnDimension->getDefaultXfIndex()
                );
            }

            $objWriter->endElement();
        }
    }

    /**
     * Write rows of the specified sheet.
     *
     * @param XMLWriter $objWriter
     * @param Worksheet $sheet
     */
    private function writeRows(XMLWriter $objWriter, Worksheet $sheet)
    {
        $numberRowsRepeated = self::NUMBER_ROWS_REPEATED_MAX;
        $span_row = 0;
        $rows = $sheet->getRowIterator();
        $lastRowFxIndex = null;
        while ($rows->valid()) {
            --$numberRowsRepeated;
            $row = $rows->current();
            if ($row->getCellIterator()->valid()) {
                if ($span_row) {
                    $objWriter->startElement('table:table-row');
                    if (null !== $lastRowFxIndex) {
                        $objWriter->writeAttribute('table:style-name', self::ROW_STYLE_PREFIX . $lastRowFxIndex);
                    }
                    if ($span_row > 1) {
                        $objWriter->writeAttribute('table:number-rows-repeated', $span_row);
                    }
                    $objWriter->startElement('table:table-cell');
                    $objWriter->writeAttribute('table:number-columns-repeated', self::NUMBER_COLS_REPEATED_MAX);
                    $objWriter->endElement();
                    $objWriter->endElement();
                    $span_row = 0;
                }
                $objWriter->startElement('table:table-row');

                $rowDimension = $sheet->getRowDimension($row->getRowIndex());
                if (null !== ($lastRowFxIndex = $rowDimension->getRowFxIndex())) {
                    $objWriter->writeAttribute('table:style-name', self::ROW_STYLE_PREFIX . $lastRowFxIndex);
                }

                $this->writeCells($objWriter, $row);
                $objWriter->endElement();
            } else {
                ++$span_row;
            }
            $rows->next();
        }
    }

    /**
     * Write cells of the specified row.
     *
     * @param XMLWriter $objWriter
     * @param Row       $row
     *
     * @throws Exception
     */
    private function writeCells(XMLWriter $objWriter, Row $row)
    {
        $numberColsRepeated = self::NUMBER_COLS_REPEATED_MAX;
        $prevColumn = -1;
        $cells = $row->getCellIterator();
        while ($cells->valid()) {
            /**
             * @var \PhpOffice\PhpSpreadsheet\Cell\Cell
             */
            $cell = $cells->current();
            $column = Coordinate::columnIndexFromString($cell->getColumn()) - 1;

            $this->writeCellSpan($objWriter, $column, $prevColumn);
            $objWriter->startElement('table:table-cell');
            $this->writeCellMerge($objWriter, $cell);

            // Style XF
            $style = $cell->getXfIndex();
            if ($style !== null) {
                $objWriter->writeAttribute('table:style-name', self::CELL_STYLE_PREFIX . $style);
            }

            switch ($cell->getDataType()) {
                case DataType::TYPE_BOOL:
                    $objWriter->writeAttribute('office:value-type', 'boolean');
                    $objWriter->writeAttribute('office:value', $cell->getValue());
                    $objWriter->writeElement('text:p', $cell->getValue());

                    break;
                case DataType::TYPE_ERROR:
                    throw new Exception('Writing of error not implemented yet.');

                    break;
                case DataType::TYPE_FORMULA:
                    $formulaValue = $cell->getValue();
                    if ($this->getParentWriter()->getPreCalculateFormulas()) {
                        try {
                            $formulaValue = $cell->getCalculatedValue();
                        } catch (Exception $e) {
                            // don't do anything
                        }
                    }
                    $objWriter->writeAttribute('table:formula', 'of:' . $cell->getValue());
                    if (is_numeric($formulaValue)) {
                        $objWriter->writeAttribute('office:value-type', 'float');
                    } else {
                        $objWriter->writeAttribute('office:value-type', 'string');
                    }
                    $objWriter->writeAttribute('office:value', $formulaValue);
                    $objWriter->writeElement('text:p', $formulaValue);

                    break;
                case DataType::TYPE_INLINE:
                    throw new Exception('Writing of inline not implemented yet.');

                    break;
                case DataType::TYPE_NUMERIC:
                    $objWriter->writeAttribute('office:value-type', 'float');
                    $objWriter->writeAttribute('office:value', $cell->getValue());
                    $objWriter->writeElement('text:p', $cell->getValue());

                    break;
                case DataType::TYPE_STRING:
                    $objWriter->writeAttribute('office:value-type', 'string');
                    $objWriter->writeElement('text:p', $cell->getValue());

                    break;
            }
            Comment::write($objWriter, $cell);
            $objWriter->endElement();
            $prevColumn = $column;
            $cells->next();
        }
        $numberColsRepeated = $numberColsRepeated - $prevColumn - 1;
        if ($numberColsRepeated > 0) {
            if ($numberColsRepeated > 1) {
                $objWriter->startElement('table:table-cell');
                $objWriter->writeAttribute('table:number-columns-repeated', $numberColsRepeated);
                $objWriter->endElement();
            } else {
                $objWriter->writeElement('table:table-cell');
            }
        }
    }

    /**
     * Write span.
     *
     * @param XMLWriter $objWriter
     * @param int       $curColumn
     * @param int       $prevColumn
     */
    private function writeCellSpan(XMLWriter $objWriter, $curColumn, $prevColumn)
    {
        $diff = $curColumn - $prevColumn - 1;
        if (1 === $diff) {
            $objWriter->writeElement('table:table-cell');
        } elseif ($diff > 1) {
            $objWriter->startElement('table:table-cell');
            $objWriter->writeAttribute('table:number-columns-repeated', $diff);
            $objWriter->endElement();
        }
    }

    /**
     * Write XF cell styles.
     *
     * @param XMLWriter   $writer
     * @param Spreadsheet $spreadsheet
     */
    private function writeXfStyles(XMLWriter $writer, Spreadsheet $spreadsheet)
    {
        foreach ($spreadsheet->additionalStyleNodes as $elem) {
            $writer->writeDomElement($elem);
        }

        foreach ($spreadsheet->getColumnXfCollection() as $style) {
            $writer->startElement('style:style');
            $writer->writeAttribute('style:name', self::COLUMN_STYLE_PREFIX . $style->getIndex());
            $writer->writeAttribute('style:family', 'table-column');
            $aData = $style->getAdditionalData();
            if (isset($aData['style:data-style-name'])) {
                $writer->writeAttribute('style:data-style-name', $aData['style:data-style-name']);
                unset($aData['style:data-style-name']);
            }

            foreach ($aData as $element => $data) {
                $writer->startElement($element);
                foreach ($data as $attr => $val) {
                    $writer->writeAttribute($attr, $val);
                }
                $writer->endElement();
            }/*
            $writer->startElement('style:table-column-properties');

            $alignment = $style->getAlignment();
            if (!empty($alignment->getBreakBefore())) {
                $writer->writeAttribute('fo:break-before', $alignment->getBreakBefore());
            }

            if (!empty($alignment->getColumnWidth())) {
                $writer->writeAttribute('style:column-width', $alignment->getColumnWidth());
            }

            $writer->endElement();*/
            $writer->endElement(); // Close style:style
        }

        foreach ($spreadsheet->getTableXfCollection() as $style) {
            $writer->startElement('style:style');
            $writer->writeAttribute('style:name', self::TABLE_STYLE_PREFIX . $style->getIndex());
            $writer->writeAttribute('style:family', 'table');
            if ($style->getFill()->getFillType() !== \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE) {
                $writer->writeAttribute('style:master-page-name', $style->getFill()->getFillType());
            }
            $aData = $style->getAdditionalData();
            if (isset($aData['style:data-style-name'])) {
                $writer->writeAttribute('style:data-style-name', $aData['style:data-style-name']);
                unset($aData['style:data-style-name']);
            }

            foreach ($aData as $element => $data) {
                $writer->startElement($element);
                foreach ($data as $attr => $val) {
                    $writer->writeAttribute($attr, $val);
                }
                $writer->endElement();
            }
            $writer->endElement(); // Close style:style
        }

        foreach ($spreadsheet->getRowXfCollection() as $style) {
            $writer->startElement('style:style');
            $writer->writeAttribute('style:name', self::ROW_STYLE_PREFIX . $style->getIndex());
            $writer->writeAttribute('style:family', 'table-row');
            $aData = $style->getAdditionalData();
            if (isset($aData['style:data-style-name'])) {
                $writer->writeAttribute('style:data-style-name', $aData['style:data-style-name']);
                unset($aData['style:data-style-name']);
            }

            foreach ($aData as $element => $data) {
                $writer->startElement($element);
                foreach ($data as $attr => $val) {
                    $writer->writeAttribute($attr, $val);
                }
                $writer->endElement();
            }/*

            $writer->startElement('style:table-row-properties');
            $alignment = $style->getAlignment();
            if (!empty($alignment->getBreakBefore())) {
                $writer->writeAttribute('fo:break-before', $alignment->getBreakBefore());
            }

            if (!empty($alignment->getRowHeight())) {
                $writer->writeAttribute('style:row-height', $alignment->getRowHeight());
            }

            if (null !== $alignment->getUseOptimalRowHeight()) {
                $writer->writeAttribute('style:use-optimal-row-height', $alignment->getUseOptimalRowHeight() ? 'true'
                    : 'false');
            }

            $writer->endElement();*/
            $writer->endElement(); // Close style:style
        }

        foreach ($spreadsheet->getCellXfCollection() as $style) {
            $writer->startElement('style:style');
            $writer->writeAttribute('style:name', self::CELL_STYLE_PREFIX . $style->getIndex());
            $writer->writeAttribute('style:family', 'table-cell');
            $writer->writeAttribute('style:parent-style-name', 'Default');
            $aData = $style->getAdditionalData();
            if (isset($aData['style:data-style-name'])) {
                $writer->writeAttribute('style:data-style-name', $aData['style:data-style-name']);
                unset($aData['style:data-style-name']);
            }

            if (is_array($aData)) {
                foreach ($aData as $element => $data) {
                    $writer->startElement($element);
                    foreach ($data as $attr => $val) {
                        $writer->writeAttribute($attr, $val);
                    }
                    $writer->endElement();
                }
            } else {
                // style:text-properties

                // Font
                $writer->startElement('style:text-properties');

                $font = $style->getFont();

                if ($font->getBold()) {
                    $writer->writeAttribute('fo:font-weight', 'bold');
                    $writer->writeAttribute('style:font-weight-complex', 'bold');
                    $writer->writeAttribute('style:font-weight-asian', 'bold');
                }

                if ($font->getItalic()) {
                    $writer->writeAttribute('fo:font-style', 'italic');
                }

                if ($color = $font->getColor()) {
                    $writer->writeAttribute('fo:color', sprintf('#%s', $color->getRGB()));
                }

                if ($family = $font->getName()) {
                    $writer->writeAttribute('style:font-family', $family);
                }

                if ($size = $font->getSize()) {
                    $writer->writeAttribute('fo:font-size', sprintf('%.1Fpt', $size));
                }

                if ($font->getUnderline() && $font->getUnderline() != Font::UNDERLINE_NONE) {
                    $writer->writeAttribute('style:text-underline-style', 'solid');
                    $writer->writeAttribute('style:text-underline-width', 'auto');
                    $writer->writeAttribute('style:text-underline-color', 'font-color');

                    switch ($font->getUnderline()) {
                        case Font::UNDERLINE_DOUBLE:
                            $writer->writeAttribute('style:text-underline-type', 'double');

                            break;
                        case Font::UNDERLINE_SINGLE:
                            $writer->writeAttribute('style:text-underline-type', 'single');

                            break;
                    }
                }

                $writer->endElement(); // Close style:text-properties

                // style:table-cell-properties

                $writer->startElement('style:table-cell-properties');
                $writer->writeAttribute('style:rotation-align', 'none');

                // Fill
                if ($fill = $style->getFill()) {
                    switch ($fill->getFillType()) {
                        case Fill::FILL_SOLID:
                            $writer->writeAttribute(
                                'fo:background-color',
                                sprintf(
                                    '#%s',
                                    strtolower($fill->getStartColor()->getRGB())
                                )
                            );

                            break;
                        case Fill::FILL_GRADIENT_LINEAR:
                        case Fill::FILL_GRADIENT_PATH:
                            /// TODO :: To be implemented
                            break;
                        case Fill::FILL_NONE:
                        default:
                    }
                }

                // Borders
                if ($borders = $style->getBorders()) {
                    foreach (['getBottom' => 'bottom', 'getLeft' => 'left', 'getTop' => 'top', 'getRight' => 'right'] as $func => $align) {
                        // @var Border
                        if (($border = $borders->{$func}()) && $border->getBorderStyle() !== Border::BORDER_NONE) {
                            $writer->writeAttribute(
                                'fo:border-' . $align,
                                $border->getSize() . 'pt ' .
                                $border->getBorderStyle() . ' #' . $border->getColor()->getARGB()
                            );
                        }
                    }
                }

                $writer->endElement(); // Close style:table-cell-properties
            }
            // End

            $writer->endElement(); // Close style:style
        }
    }

    /**
     * Write attributes for merged cell.
     *
     * @param XMLWriter $objWriter
     * @param Cell      $cell
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function writeCellMerge(XMLWriter $objWriter, Cell $cell)
    {
        if (!$cell->isMergeRangeValueCell()) {
            return;
        }

        $mergeRange = Coordinate::splitRange($cell->getMergeRange());
        [$startCell, $endCell] = $mergeRange[0];
        $start = Coordinate::coordinateFromString($startCell);
        $end = Coordinate::coordinateFromString($endCell);
        $columnSpan = Coordinate::columnIndexFromString($end[0]) - Coordinate::columnIndexFromString($start[0]) + 1;
        $rowSpan = $end[1] - $start[1] + 1;

        $objWriter->writeAttribute('table:number-columns-spanned', $columnSpan);
        $objWriter->writeAttribute('table:number-rows-spanned', $rowSpan);
    }
}

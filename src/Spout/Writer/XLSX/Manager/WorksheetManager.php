<?php

namespace Box\Spout3\Writer\XLSX\Manager;

use Box\Spout3\Common\Entity\Cell;
use Box\Spout3\Common\Entity\Row;
use Box\Spout3\Common\Entity\Style\Style;
use Box\Spout3\Common\Exception\InvalidArgumentException;
use Box\Spout3\Common\Exception\IOException;
use Box\Spout3\Common\Helper\Escaper\XLSX as XLSXEscaper;
use Box\Spout3\Common\Helper\StringHelper;
use Box\Spout3\Common\Manager\OptionsManagerInterface;
use Box\Spout3\Writer\Common\Creator\InternalEntityFactory;
use Box\Spout3\Writer\Common\Entity\Options;
use Box\Spout3\Writer\Common\Entity\Worksheet;
use Box\Spout3\Writer\Common\Helper\CellHelper;
use Box\Spout3\Writer\Common\Manager\RowManager;
use Box\Spout3\Writer\Common\Manager\Style\StyleMerger;
use Box\Spout3\Writer\Common\Manager\WorksheetManagerInterface;
use Box\Spout3\Writer\XLSX\Manager\Style\StyleManager;

/**
 * Class WorksheetManager
 * XLSX worksheet manager, providing the interfaces to work with XLSX worksheets.
 */
class WorksheetManager implements WorksheetManagerInterface
{
    /**
     * Maximum number of characters a cell can contain
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-16c69c74-3d6a-4aaf-ba35-e6eb276e8eaa [Excel 2007]
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3 [Excel 2010]
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-ca36e2dc-1f09-4620-b726-67c00b05040f [Excel 2013/2016]
     */
    const MAX_CHARACTERS_PER_CELL = 32767;

    const SHEET_XML_FILE_HEADER = <<<'EOD'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
EOD;

    /** @var bool Whether inline or shared strings should be used */
    protected $shouldUseInlineStrings;

    /** @var RowManager Manages rows */
    private $rowManager;

    /** @var StyleManager Manages styles */
    private $styleManager;

    /** @var StyleMerger Helper to merge styles together */
    private $styleMerger;

    /** @var SharedStringsManager Helper to write shared strings */
    private $sharedStringsManager;

    /** @var XLSXEscaper Strings escaper */
    private $stringsEscaper;

    /** @var StringHelper String helper */
    private $stringHelper;

    /** @var InternalEntityFactory Factory to create entities */
    private $entityFactory;

    /** @var int The highest number of columns in any row. */
    private $maxColumns = 0;

    /** @var bool Whether to merge cell styles with row styles, or just use the cell style unaltered if one is provided. */
    protected $mergeCellStyles = false;

    /** @var array All tooltips associated with cells in this worksheet. */
    public $tooltips = [];

    /** @var array An array of column options */
    public $options = [];

    /** @var resource A filepointer to the currently opened sheet. */
    private $active_sheet_file_pointer;

    private $active_sheet_rels_file_pointer;

    /** @var resource A filepointer to the currently opened drawings. */
    private $active_drawing_file_pointer;

    private $active_drawing_rels_file_pointer;

    private $current_image_offset = 0;

    private $active_sheet;

    private $queued_images = [];

    /**
     * WorksheetManager constructor.
     *
     * @param OptionsManagerInterface $optionsManager
     * @param RowManager $rowManager
     * @param StyleManager $styleManager
     * @param StyleMerger $styleMerger
     * @param SharedStringsManager $sharedStringsManager
     * @param XLSXEscaper $stringsEscaper
     * @param StringHelper $stringHelper
     * @param InternalEntityFactory $entityFactory
     */
    public function __construct(
        OptionsManagerInterface $optionsManager,
        RowManager $rowManager,
        StyleManager $styleManager,
        StyleMerger $styleMerger,
        SharedStringsManager $sharedStringsManager,
        XLSXEscaper $stringsEscaper,
        StringHelper $stringHelper,
        InternalEntityFactory $entityFactory
    ) {
        $this->shouldUseInlineStrings = $optionsManager->getOption(Options::SHOULD_USE_INLINE_STRINGS);
        $this->rowManager = $rowManager;
        $this->styleManager = $styleManager;
        $this->styleMerger = $styleMerger;
        $this->sharedStringsManager = $sharedStringsManager;
        $this->stringsEscaper = $stringsEscaper;
        $this->stringHelper = $stringHelper;
        $this->entityFactory = $entityFactory;
    }

    /**
     * @return SharedStringsManager
     */
    public function getSharedStringsManager()
    {
        return $this->sharedStringsManager;
    }

    /**
     * {@inheritdoc}
     * TODO: sheetView should be boolean, enabled by default; cols should be exposed via createSheet, I recommend connecting using headers
     */
    public function startSheet(Worksheet $worksheet, $options = [])
    {

        // the many hacks in this fork make this somewhat necessary
        if ($this->active_sheet) {
            $this->close($this->active_sheet);
        }

        $this->active_sheet = $worksheet;



        // Open the file pointer for the sheet
        $this->active_sheet_file_pointer = \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fopen_buffered($worksheet->getFilePath(), 'w');

        $this->throwIfSheetFilePointerIsNotAvailable($this->active_sheet_file_pointer);

        // Begin writing the sheet
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, self::SHEET_XML_FILE_HEADER);


        // Open the file pointer for the sheet rels
        $this->active_sheet_rels_file_pointer = \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fopen_buffered(
            dirname($worksheet->getFilePath())
                . '/_rels/'
                . basename($worksheet->getFilePath())
                . '.rels'
            , 'w'
        );

        $this->throwIfSheetFilePointerIsNotAvailable($this->active_sheet_rels_file_pointer);


        // Derive the drawings file name from the sheet name
        $drawings_name = str_replace('sheet', 'drawing', basename($worksheet->getFilePath()));


        // Link the drawings into the sheet rels
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered(
            $this->active_sheet_rels_file_pointer,
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                . "\n"
                . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/' . $drawings_name . '"/></Relationships>'
        );


        // Open the drawings file
        $this->active_drawing_file_pointer = \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fopen_buffered(
            dirname($worksheet->getFilePath())
                . '/../drawings/'
                . $drawings_name,
            'w'
        );

        $this->throwIfSheetFilePointerIsNotAvailable($this->active_drawing_file_pointer);

        // Start the drawings file
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered(
            $this->active_drawing_file_pointer,
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                . "\n"
                . '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        );


        // Open the drawings rels file (this links drawing XML objects to physical image files)
        $this->active_drawing_rels_file_pointer = \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fopen_buffered(
            dirname($worksheet->getFilePath())
                . '/../drawings/_rels/'
                . $drawings_name
                . '.rels',
            'w'
        );

        $this->throwIfSheetFilePointerIsNotAvailable($this->active_drawing_rels_file_pointer);


        // Start the drawings rels file
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered(
            $this->active_drawing_rels_file_pointer,
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                . "\n"
                . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        );


        // Let the worksheet know what its file pointer is
        $worksheet->setFilePointer($this->active_sheet_file_pointer);


        // Let the worksheet know what options are applied to it
        foreach ($options AS $option => $value) {
            $worksheet->setOption($option, $value);
        }



        // Apply freeze pane, if applicable
        if (!empty($this->active_sheet->getOption('freeze_pane'))) {
            $xpos = $this->active_sheet->getOption('freeze_pane')[0] ?? 0;
            $ypos = $this->active_sheet->getOption('freeze_pane')[1] ?? 0;

            $topRight = self::getCellOffset($xpos + 1, $ypos);
            $bottomLeft = self::getCellOffset(1, $ypos);
            $bottomRight = self::getCellOffset($xpos + 1, $ypos);

            // Disables copy/paste, I don't know why
            \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '
                <sheetViews>
                    <sheetView tabSelected="1" workbookViewId="0">
                        <pane xSplit="' . $xpos . '" ySplit="' . $ypos . '" topLeftCell="' . $bottomRight . '" activePane="bottomRight" state="frozen" />
                        <selection pane="topRight" activeCell="' . $topRight . '" sqref="' . $topRight . '"  />
                        <selection pane="bottomLeft" activeCell="' . $bottomLeft . '" sqref="' . $bottomLeft . '"  />
                        <selection pane="bottomRight" activeCell="' . $bottomRight . '" sqref="' . $bottomRight . '" />
                    </sheetView>
                </sheetViews>
            ');
        }



        // Apply column widths, if applicable
        if (!empty($this->active_sheet->getOption('column_widths'))) {
            \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '<cols>');

            foreach ($this->active_sheet->getOption('column_widths') AS $i => $width) {
                if ($width) {
                    \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '<col min="' . ($i + 1) . '" max="' . ($i + 1) . '" width="' . $width . '" customWidth="1" />');
                }
            }

            \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '</cols>');
        }



        // Begin the <sheetData> XML block
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '<sheetData>');
    }

    /**
     * Checks if the sheet has been sucessfully created. Throws an exception if not.
     *
     * @param bool|resource $sheetFilePointer Pointer to the sheet data file or FALSE if unable to open the file
     * @throws IOException If the sheet data file cannot be opened for writing
     * @return void
     */
    private function throwIfSheetFilePointerIsNotAvailable($sheetFilePointer)
    {
        if (!$sheetFilePointer) {
            throw new IOException('Unable to open sheet for writing.');
        }
    }

    /**
     * {@inheritdoc}
     */
    public function addRow(Worksheet $worksheet, Row $row)
    {
        if (!$this->rowManager->isEmpty($row)) {
            $this->addNonEmptyRow($worksheet, $row);
        }

        $worksheet->setLastWrittenRowIndex($worksheet->getLastWrittenRowIndex() + 1);
    }

    /**
     * Adds non empty row to the worksheet.
     *
     * @param Worksheet $worksheet The worksheet to add the row to
     * @param Row $row The row to be written
     * @throws IOException If the data cannot be written
     * @throws InvalidArgumentException If a cell value's type is not supported
     * @return void
     */
    private function addNonEmptyRow(Worksheet $worksheet, Row $row)
    {
        $cellIndex = 0;
        $rowStyle = $row->getStyle();
        $rowIndex = $worksheet->getLastWrittenRowIndex() + 1;
        $numCells = $row->getNumCells();
        $this->maxColumns = max($this->maxColumns, $numCells);

        $rowXML = '<row r="' . $rowIndex . '" spans="1:' . $numCells . '"';

        // if the row contains an image, set the height to 100
        foreach ($row->getCells() AS $cell) {
            if ($cell->isImage()) {
                $rowXML .= ' ht="100" customHeight="1"';
                break;
            }
        }

        $rowXML .= '>';

        foreach ($row->getCells() as $cell) {
            $rowXML .= $this->applyStyleAndGetCellXML($cell, $rowStyle, $rowIndex, $cellIndex);

            if ($cell->getTooltip()) {
                $this->tooltips[self::getCellOffset($rowIndex, $cellIndex)] = $cell->getTooltip();
            }

            $cellIndex++;
        }

        $rowXML .= '</row>';

        $wasWriteSuccessful = \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($worksheet->getFilePointer(), $rowXML);
        if ($wasWriteSuccessful === false) {
            throw new IOException("Unable to write data in {$worksheet->getFilePath()}");
        }
    }

    /**
     * Applies styles to the given style, merging the cell's style with its row's style
     * Then builds and returns xml for the cell.
     *
     * @param Cell $cell
     * @param Style $rowStyle
     * @param int $rowIndex
     * @param int $cellIndex
     * @throws InvalidArgumentException If the given value cannot be processed
     * @return string
     */
    private function applyStyleAndGetCellXML(Cell $cell, ?Style $rowStyle, $rowIndex, $cellIndex)
    {

        // Apply row and extra styles
        // Perform a full (but slower) merge of cell and row styles if enabled.
        if ($rowStyle && $cell->getStyle()) {
            $mergedCellAndRowStyle = $this->styleMerger->merge($cell->getStyle(), $rowStyle);
            $cell->setStyle($mergedCellAndRowStyle);
            $newCellStyle = $this->styleManager->applyExtraStylesIfNeeded($cell);
            $registeredStyle = $this->styleManager->registerStyle($newCellStyle)->getId();
        }
        elseif ($rowStyle) {
            $registeredStyle = $this->styleManager->registerStyle($rowStyle)->getId();
        }
        elseif ($cell->getStyle()) {
            $registeredStyle = $this->styleManager->registerStyle($cell->getStyle())->getId();
        }
        else {
            $registeredStyle = null;
        }

        return $this->getCellXML($rowIndex, $cellIndex, $cell, $registeredStyle);
    }

    /**
     * Builds and returns xml for a single cell.
     *
     * @param int $rowIndex
     * @param int $cellNumber
     * @param Cell $cell
     * @param int $styleId
     * @throws InvalidArgumentException If the given value cannot be processed
     * @return string
     */
    private function getCellXML($rowIndex, $cellNumber, Cell $cell, $styleId)
    {

        // special case: don't do anything if cell is empty and unstyled
        if ($cell->isEmpty() && (!$styleId || !$this->styleManager->shouldApplyStyleOnEmptyCell($styleId))) {
            return '';
        }

        // special case: if cell is image, write it
        if ($cell->isImage()) {
            if (!empty($cell->getValue())) {
                // Add the image to the queue of images to load
                $this->queued_images[] = [$rowIndex, $cellNumber, $cell->getValue()];
            }

            // We aren't rendering any cell data.
            return '';
        }


        $columnIndex = CellHelper::getCellIndexFromColumnIndex($cellNumber);
        $cellXML = '<c r="' . $columnIndex . $rowIndex . '"';

        if ($styleId) {
            $cellXML .= ' s="' . $styleId . '"';
        }

        if ($cell->isString()) {
            $cellXML .= $this->getCellXMLFragmentForNonEmptyString($cell->getValue());
        } elseif ($cell->isBoolean()) {
            $cellXML .= ' t="b"><v>' . (int) ($cell->getValue()) . '</v></c>';
        } elseif ($cell->isNumeric()) {
            $cellXML .= '><v>' . $cell->getValue() . '</v></c>';
        } elseif ($cell->isEmpty()) {
            $cellXML .= '/>';
        } else {
            throw new InvalidArgumentException('Trying to add a value with an unsupported type: ' . \gettype($cell->getValue()));
        }

        return $cellXML;

    }

    /**
     * Returns the XML fragment for a cell containing a non empty string
     *
     * @param string $cellValue The cell value
     * @throws InvalidArgumentException If the string exceeds the maximum number of characters allowed per cell
     * @return string The XML fragment representing the cell
     */
    private function getCellXMLFragmentForNonEmptyString($cellValue)
    {
        if ($this->stringHelper->getStringLength($cellValue) > self::MAX_CHARACTERS_PER_CELL) {
            throw new InvalidArgumentException('Trying to add a value that exceeds the maximum number of characters allowed in a cell (32,767)');
        }

        if ($this->shouldUseInlineStrings) {
            $cellXMLFragment = ' t="inlineStr"><is><t>' . $this->stringsEscaper->escape($cellValue) . '</t></is></c>';
        } else {
            $sharedStringId = $this->sharedStringsManager->writeString($cellValue);
            $cellXMLFragment = ' t="s"><v>' . $sharedStringId . '</v></c>';
        }

        return $cellXMLFragment;
    }


    /**
     * {@inheritdoc}
     */
    public function close(Worksheet $worksheet)
    {

        if (!is_resource($this->active_sheet_file_pointer)) {
            return;
        }

        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '</sheetData>');

        if ($worksheet->getOption('filter') && $worksheet->getMaxNumColumns()) {

            \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '<autoFilter ref="A1:ZZ1"/>');

            if ($this->tooltips) {
                \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered(
                    $this->active_sheet_file_pointer,
                    '<dataValidations count="' . \count($this->tooltips) . '">'
                        . array_map(function($tooltip, $cell_ref) {
                            return '<dataValidation allowBlank="1" showInputMessage="1" showErrorMessage="1" prompt="' . $tooltip . '" sqref="' . $cell_ref . '" />';
                        }, $this->tooltips, array_keys($this->tooltips))
                        . '</dataValidations>'
                );
            }
        }


        // Resolve every image that is being inserted
        $client = new \GuzzleHttp\Client();

        $pool = new \GuzzleHttp\Pool(
            $client,
            array_map(
                function($image) {
                    // We proxy this request to enable caching. We override the default ttl, 12 hours, with a 7 day ttl -- for the purposes of this kind-of thing, we really don't care that much if the image is out-of-date.
                    return new \GuzzleHttp\Psr7\Request('GET', 'https://hamr.mmm.com/proxy?ttl=604800&url=' . $image);
                },
                array_column($this->queued_images, 2)
            ),
            [
                'concurrency' => 20,
                'rejected' => function ($reason, $index) {
                    // do nothing
                },
                'fulfilled' => function (\GuzzleHttp\Psr7\Response $response, $index) {

                    $contents = $response->getBody()->getContents();

                    $size = getimagesizefromstring($contents);
                    $dpi = hexdec(substr(bin2hex(substr($contents,14,4)),0,4));

                    [$rowIndex, $cellNumber, $image] = $this->queued_images[$index];

                    $image_name = sha1($image) . '-' . basename($image);

                    // Write the anchor to the drawings file
                    \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_drawing_file_pointer, '
                        <xdr:twoCellAnchor>
                            <xdr:from>
                                <xdr:col>' . $cellNumber . '</xdr:col>
                                <xdr:colOff>0</xdr:colOff>
                                <xdr:row>' . ($rowIndex - 1) . '</xdr:row>
                                <xdr:rowOff>0</xdr:rowOff>
                            </xdr:from>
                            <xdr:to>
                                <xdr:col>' . $cellNumber . '</xdr:col>
                                <xdr:colOff>' . $size[0] * 12000 . '</xdr:colOff>
                                <xdr:row>' . ($rowIndex - 1) . '</xdr:row>
                                <xdr:rowOff>1270000</xdr:rowOff>
                            </xdr:to>
                            <xdr:pic>
                                <xdr:nvPicPr>
                                    <xdr:cNvPr id="' . ++$this->current_image_offset . '" name="" descr="">
                                    </xdr:cNvPr>
                                    <xdr:cNvPicPr>
                                        <a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
                                    </xdr:cNvPicPr>
                                </xdr:nvPicPr>
                                <xdr:blipFill>
                                    <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' . $this->current_image_offset . '">
                                    </a:blip>
                                    <a:srcRect/>
                                    <a:stretch>
                                        <a:fillRect/>
                                    </a:stretch>
                                </xdr:blipFill>
                                <xdr:spPr bwMode="auto">
                                    <a:xfrm>
                                    </a:xfrm>
                                    <a:prstGeom prst="rect">
                                        <a:avLst/>
                                    </a:prstGeom>
                                    <a:noFill/>
                                </xdr:spPr>
                            </xdr:pic>
                            <xdr:clientData/>
                        </xdr:twoCellAnchor>
                    ');

                    // Write the relationship to the rels file
                    \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_drawing_rels_file_pointer, '<Relationship Id="rId' . $this->current_image_offset . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' . $image_name.'" />');

                    file_put_contents(
                        \dirname($this->active_sheet->getFilePath())
                            . '/../media/'
                            . $image_name,
                        $contents
                    );
                    // this is delivered each successful response
                },
            ]
        );

        // Initiate the transfers and create a promise
        $promise = $pool->promise();

        // Force the pool of requests to complete.
        $promise->wait();


        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '<drawing r:id="rId1"/>');
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered($this->active_sheet_file_pointer, '</worksheet>');

        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered(
            $this->active_drawing_file_pointer,
            '</xdr:wsDr>'
        );

        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fwrite_buffered(
            $this->active_drawing_rels_file_pointer,
            '</Relationships>'
        );

        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fclose_buffered($this->active_sheet_file_pointer);
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fclose_buffered($this->active_sheet_rels_file_pointer);
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fclose_buffered($this->active_drawing_file_pointer);
        \Box\Spout3\Common\Helper\GlobalFunctionsHelper::fclose_buffered($this->active_drawing_rels_file_pointer);

    }


    public static function getCellOffset($x, $y)
    {
        return chr(64 + $x)
            . ($y + 1);
    }
}

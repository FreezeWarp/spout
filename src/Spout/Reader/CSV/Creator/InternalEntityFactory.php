<?php

namespace Box\Spout3\Reader\CSV\Creator;

use Box\Spout3\Common\Creator\HelperFactory;
use Box\Spout3\Common\Entity\Cell;
use Box\Spout3\Common\Entity\Row;
use Box\Spout3\Common\Helper\GlobalFunctionsHelper;
use Box\Spout3\Common\Manager\OptionsManagerInterface;
use Box\Spout3\Reader\Common\Creator\InternalEntityFactoryInterface;
use Box\Spout3\Reader\CSV\RowIterator;
use Box\Spout3\Reader\CSV\Sheet;
use Box\Spout3\Reader\CSV\SheetIterator;

/**
 * Class EntityFactory
 * Factory to create entities
 */
class InternalEntityFactory implements InternalEntityFactoryInterface
{
    /** @var HelperFactory */
    private $helperFactory;

    /**
     * @param HelperFactory $helperFactory
     */
    public function __construct(HelperFactory $helperFactory)
    {
        $this->helperFactory = $helperFactory;
    }

    /**
     * @param resource $filePointer Pointer to the CSV file to read
     * @param OptionsManagerInterface $optionsManager
     * @param GlobalFunctionsHelper $globalFunctionsHelper
     * @return SheetIterator
     */
    public function createSheetIterator($filePointer, $optionsManager, $globalFunctionsHelper)
    {
        $rowIterator = $this->createRowIterator($filePointer, $optionsManager, $globalFunctionsHelper);
        $sheet = $this->createSheet($rowIterator);

        return new SheetIterator($sheet);
    }

    /**
     * @param RowIterator $rowIterator
     * @return Sheet
     */
    private function createSheet($rowIterator)
    {
        return new Sheet($rowIterator);
    }

    /**
     * @param resource $filePointer Pointer to the CSV file to read
     * @param OptionsManagerInterface $optionsManager
     * @param GlobalFunctionsHelper $globalFunctionsHelper
     * @return RowIterator
     */
    private function createRowIterator($filePointer, $optionsManager, $globalFunctionsHelper)
    {
        $encodingHelper = $this->helperFactory->createEncodingHelper($globalFunctionsHelper);

        return new RowIterator($filePointer, $optionsManager, $encodingHelper, $this, $globalFunctionsHelper);
    }

    /**
     * @param Cell[] $cells
     * @return Row
     */
    public function createRow(array $cells = [])
    {
        return new Row($cells, null);
    }

    /**
     * @param mixed $cellValue
     * @return Cell
     */
    public function createCell($cellValue)
    {
        return new Cell($cellValue);
    }

    /**
     * @param array $cellValues
     * @return Row
     */
    public function createRowFromArray(array $cellValues = [])
    {
        $cells = array_map(function ($cellValue) {
            return $this->createCell($cellValue);
        }, $cellValues);

        return $this->createRow($cells);
    }
}

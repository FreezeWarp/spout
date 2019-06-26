<?php

namespace Box\Spout3\Reader\Common\Creator;

use Box\Spout3\Common\Entity\Cell;
use Box\Spout3\Common\Entity\Row;

/**
 * Interface EntityFactoryInterface
 */
interface InternalEntityFactoryInterface
{
    /**
     * @param Cell[] $cells
     * @return Row
     */
    public function createRow(array $cells = []);

    /**
     * @param mixed $cellValue
     * @return Cell
     */
    public function createCell($cellValue);
}

<?php

namespace Box\Spout3\Writer\Common\Manager;

use Box\Spout3\Common\Entity\Row;

class RowManager
{
    /**
     * Detect whether a row is considered empty.
     * An empty row has all of its cells empty.
     *
     * @param Row $row
     * @return bool
     */
    public function isEmpty(Row $row)
    {
        foreach ($row->getCells() as $cell) {
            if (!$cell->isEmpty()) {
                return false;
            }
        }

        return true;
    }
}

<?php

namespace Box\Spout3\Reader\ODS\Creator;

use Box\Spout3\Reader\ODS\Helper\CellValueFormatter;
use Box\Spout3\Reader\ODS\Helper\SettingsHelper;

/**
 * Class EntityFactory
 * Factory to create helpers
 */
class HelperFactory extends \Box\Spout3\Common\Creator\HelperFactory
{
    /**
     * @param bool $shouldFormatDates Whether date/time values should be returned as PHP objects or be formatted as strings
     * @return CellValueFormatter
     */
    public function createCellValueFormatter($shouldFormatDates)
    {
        $escaper = $this->createStringsEscaper();

        return new CellValueFormatter($shouldFormatDates, $escaper);
    }

    /**
     * @param InternalEntityFactory $entityFactory
     * @return SettingsHelper
     */
    public function createSettingsHelper($entityFactory)
    {
        return new SettingsHelper($entityFactory);
    }

    /**
     * @return \Box\Spout3\Common\Helper\Escaper\ODS
     */
    public function createStringsEscaper()
    {
        /* @noinspection PhpUnnecessaryFullyQualifiedNameInspection */
        return new \Box\Spout3\Common\Helper\Escaper\ODS();
    }
}

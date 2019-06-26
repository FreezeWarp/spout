<?php

namespace Box\Spout3\Writer\Common\Creator;

use Box\Spout3\Common\Manager\OptionsManagerInterface;
use Box\Spout3\Writer\Common\Manager\SheetManager;
use Box\Spout3\Writer\Common\Manager\WorkbookManagerInterface;

/**
 * Interface ManagerFactoryInterface
 */
interface ManagerFactoryInterface
{
    /**
     * @param OptionsManagerInterface $optionsManager
     * @return WorkbookManagerInterface
     */
    public function createWorkbookManager(OptionsManagerInterface $optionsManager);

    /**
     * @return SheetManager
     */
    public function createSheetManager();
}

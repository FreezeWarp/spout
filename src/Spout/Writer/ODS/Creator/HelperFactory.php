<?php

namespace Box\Spout3\Writer\ODS\Creator;

use Box\Spout3\Common\Helper\Escaper;
use Box\Spout3\Common\Helper\StringHelper;
use Box\Spout3\Common\Manager\OptionsManagerInterface;
use Box\Spout3\Writer\Common\Creator\InternalEntityFactory;
use Box\Spout3\Writer\Common\Entity\Options;
use Box\Spout3\Writer\Common\Helper\ZipHelper;
use Box\Spout3\Writer\ODS\Helper\FileSystemHelper;

/**
 * Class HelperFactory
 * Factory for helpers needed by the ODS Writer
 */
class HelperFactory extends \Box\Spout3\Common\Creator\HelperFactory
{
    /**
     * @param OptionsManagerInterface $optionsManager
     * @param InternalEntityFactory $entityFactory
     * @return FileSystemHelper
     */
    public function createSpecificFileSystemHelper(OptionsManagerInterface $optionsManager, InternalEntityFactory $entityFactory)
    {
        $tempFolder = $optionsManager->getOption(Options::TEMP_FOLDER);
        $zipHelper = $this->createZipHelper($entityFactory);

        return new FileSystemHelper($tempFolder, $zipHelper);
    }

    /**
     * @param $entityFactory
     * @return ZipHelper
     */
    private function createZipHelper($entityFactory)
    {
        return new ZipHelper($entityFactory);
    }

    /**
     * @return Escaper\ODS
     */
    public function createStringsEscaper()
    {
        return new Escaper\ODS();
    }

    /**
     * @return StringHelper
     */
    public function createStringHelper()
    {
        return new StringHelper();
    }
}

<?php

namespace Box\Spout3\Common\Creator;

use Box\Spout3\Common\Helper\EncodingHelper;
use Box\Spout3\Common\Helper\FileSystemHelper;
use Box\Spout3\Common\Helper\GlobalFunctionsHelper;
use Box\Spout3\Common\Helper\StringHelper;

/**
 * Class HelperFactory
 * Factory to create helpers
 */
class HelperFactory
{
    /**
     * @return GlobalFunctionsHelper
     */
    public function createGlobalFunctionsHelper()
    {
        return new GlobalFunctionsHelper();
    }

    /**
     * @param string $baseFolderPath The path of the base folder where all the I/O can occur
     * @return FileSystemHelper
     */
    public function createFileSystemHelper($baseFolderPath)
    {
        return new FileSystemHelper($baseFolderPath);
    }

    /**
     * @param GlobalFunctionsHelper $globalFunctionsHelper
     * @return EncodingHelper
     */
    public function createEncodingHelper(GlobalFunctionsHelper $globalFunctionsHelper)
    {
        return new EncodingHelper($globalFunctionsHelper);
    }

    /**
     * @return StringHelper
     */
    public function createStringHelper()
    {
        return new StringHelper();
    }
}

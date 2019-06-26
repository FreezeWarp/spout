<?php

namespace Box\Spout3\Writer\Common\Helper;

use Box\Spout3\Common\Helper\FileSystemHelperInterface;

/**
 * Class FileSystemHelperInterface
 * This interface describes helper functions to help with the file system operations
 * like files/folders creation & deletion
 */
interface FileSystemWithRootFolderHelperInterface extends FileSystemHelperInterface
{
    /**
     * Creates all the folders needed to create a spreadsheet, as well as the files that won't change.
     *
     * @throws \Box\Spout3\Common\Exception\IOException If unable to create at least one of the base folders
     * @return void
     */
    public function createBaseFilesAndFolders();

    /**
     * @return string
     */
    public function getRootFolder();
}

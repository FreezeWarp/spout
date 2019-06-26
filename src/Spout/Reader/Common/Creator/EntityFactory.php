<?php

namespace Box\Spout3\Reader\Common\Creator;

use Box\Spout3\Reader\ReaderFactory;
use Box\Spout3\Reader\ReaderInterface;

/**
 * Class EntityFactory
 * Factory to create external entities
 */
class EntityFactory
{
    /**
     * This creates an instance of the appropriate reader, given the type of the file to be read
     *
     * @param  string $readerType Type of the reader to instantiate
     * @throws \Box\Spout3\Common\Exception\UnsupportedTypeException
     * @return ReaderInterface
     */
    public static function createReader($readerType)
    {
        return (new ReaderFactory())->create($readerType);
    }
}

<?php

namespace Box\Spout3\Reader;

/**
 * Interface IteratorInterface
 */
interface IteratorInterface extends \Iterator
{
    /**
     * Cleans up what was created to iterate over the object.
     *
     * @return void
     */
    public function end();
}

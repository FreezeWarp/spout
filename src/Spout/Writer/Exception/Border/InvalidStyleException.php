<?php

namespace Box\Spout3\Writer\Exception\Border;

use Box\Spout3\Common\Entity\Style\BorderPart;
use Box\Spout3\Writer\Exception\WriterException;

class InvalidStyleException extends WriterException
{
    public function __construct($name)
    {
        $msg = '%s is not a valid style identifier for a border. Valid identifiers are: %s.';

        parent::__construct(sprintf($msg, $name, implode(',', BorderPart::getAllowedStyles())));
    }
}

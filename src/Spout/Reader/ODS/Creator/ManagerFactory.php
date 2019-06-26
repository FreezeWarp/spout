<?php

namespace Box\Spout3\Reader\ODS\Creator;

use Box\Spout3\Reader\Common\Manager\RowManager;

/**
 * Class ManagerFactory
 * Factory to create managers
 */
class ManagerFactory
{
    /**
     * @param InternalEntityFactory $entityFactory Factory to create entities
     * @return RowManager
     */
    public function createRowManager($entityFactory)
    {
        return new RowManager($entityFactory);
    }
}

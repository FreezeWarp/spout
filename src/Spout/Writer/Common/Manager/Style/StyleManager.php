<?php

namespace Box\Spout3\Writer\Common\Manager\Style;

use Box\Spout3\Common\Entity\Cell;
use Box\Spout3\Common\Entity\Style\Style;

/**
 * Class StyleManager
 * Manages styles to be applied to a cell
 */
class StyleManager implements StyleManagerInterface
{
    /** @var StyleRegistry Registry for all used styles */
    protected $styleRegistry;

    /**
     * @param StyleRegistry $styleRegistry
     */
    public function __construct(StyleRegistry $styleRegistry)
    {
        $this->styleRegistry = $styleRegistry;
    }

    /**
     * Returns the default style
     *
     * @return Style Default style
     */
    protected function getDefaultStyle()
    {
        // By construction, the default style has ID 0
        return $this->styleRegistry->getRegisteredStyles()[0];
    }

    /**
     * Registers the given style as a used style.
     * Duplicate styles won't be registered more than once.
     *
     * @param Style $style The style to be registered
     * @return Style The registered style, updated with an internal ID.
     */
    public function registerStyle($style)
    {
        return $this->styleRegistry->registerStyle($style);
    }

    /**
     * Apply additional styles if the given row needs it.
     * Typically, set "wrap text" if a cell contains a new line.
     *
     * @param Cell $cell
     * @return Style
     */
    public function applyExtraStylesIfNeeded(Cell $cell)
    {
        return $cell->getStyle();
    }
}

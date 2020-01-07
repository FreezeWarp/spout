<?php

namespace Box\Spout3\Common\Entity;

use Box\Spout3\Common\Entity\Style\Style;
use Box\Spout3\Common\Helper\CellTypeHelper;

/**
 * Class Cell
 */
class Cell
{
    /**
     * Numeric cell type (whole numbers, fractional numbers, dates)
     */
    const TYPE_NUMERIC = 0;

    /**
     * String (text) cell type
     */
    const TYPE_STRING = 1;

    /**
     * Formula cell type
     * Not used at the moment
     */
    const TYPE_FORMULA = 2;

    /**
     * Empty cell type
     */
    const TYPE_EMPTY = 3;

    /**
     * Boolean cell type
     */
    const TYPE_BOOLEAN = 4;

    /**
     * Date cell type
     */
    const TYPE_DATE = 5;

    /**
     * Error cell type
     */
    const TYPE_ERROR = 6;

    /**
     * Image cell type (requires special handling... a lot of it; assumes input is URLs)
     */
    const TYPE_IMAGE = 7;

    /**
     * Link cell type (requires special handling)
     */
    const TYPE_LINK = 8;

    /**
     * The value of this cell
     * @var mixed|null
     */
    protected $value;

    /**
     * The cell type
     * @var int|null
     */
    protected $type;

    /**
     * The cell style
     * @var Style
     */
    protected $style;

    /**
     * The cell tooltip
     * @var tooltip
     */
    protected $tooltip;

    /**
     * @param $value mixed
     * @param Style|null $style
     */
    public function __construct($value, ?Style $style = null)
    {
        $this->setValue($value);
        $this->setStyle($style);
    }

    /**
     * @param mixed|null $value
     */
    public function setValue($value)
    {
        $this->value = $value;
    }

    /**
     * @return mixed|null
     */
    public function getValue()
    {
        return $this->value;
    }

    /**
     * @param String $tooltip
     */
    public function setTooltip(?String $tooltip)
    {
        $this->tooltip = $tooltip;
    }

    /**
     * @return String
     */
    public function getTooltip(): ?String
    {
        return $this->tooltip;
    }

    /**
     * @param Style|null $style
     */
    public function setStyle(?Style $style)
    {
        $this->style = $style;
    }

    /**
     * @return Style
     */
    public function getStyle(): ?Style
    {
        return $this->style;
    }

    /**
     * @return int|null
     */
    public function getType()
    {
        if (!empty($this->type)) {
            return $this->type;
        } else {
            return $this->type = $this->detectType($this->value);
        }
    }

    /**
     * @param int $type
     */
    public function setType($type)
    {
        $this->type = $type;
    }

    /**
     * Get the current value type
     *
     * @param mixed|null $value
     * @return int
     */
    protected function detectType($value)
    {
        if (is_bool($value)) {
            return self::TYPE_BOOLEAN;
        }
        if (is_null($value) || $value === "") {
            return self::TYPE_EMPTY;
        }
        if (is_int($value) || is_float($value) || is_bool($value)) {
            return self::TYPE_NUMERIC;
        }
        if ($value instanceof \DateTime ||
            $value instanceof \DateInterval) {
            return self::TYPE_DATE;
        }
        if (is_string($value)) {
            return self::TYPE_STRING;
        }

        return self::TYPE_ERROR;
    }

    /**
     * @return bool
     */
    public function isEmpty()
    {
        return $this->type === self::TYPE_EMPTY;
    }

    /**
     * @return bool
     */
    public function isError()
    {
        return $this->type === self::TYPE_ERROR;
    }

    /**
     * @return string
     */
    public function __toString()
    {
        return (string) $this->value;
    }
}

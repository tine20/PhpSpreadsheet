<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Ods;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Style;

abstract class AbstractStyle
{
    protected $styleMap = [];

    protected $styleNS;

    protected $fontNS;

    protected $spreadsheet;

    public function __construct(Spreadsheet $spreadsheet, $styleNS, $fontNS)
    {
        $this->spreadsheet = $spreadsheet;
        $this->styleNS = $styleNS;
        $this->fontNS = $fontNS;
    }

    /**
     * @param \DOMElement $element
     *
     * @return null|Style
     */
    protected function _readFromDom(\DOMElement $element)
    {
        $styleData = new Style();

        $additionalData = [];
        /** @var \DOMElement $child */
        foreach ($element->childNodes as $child) {
            [$ns, ] = explode(':', $child->nodeName, 2);
            if ($child->hasAttributes() && 'style' === $ns) {
                $data = [];
                /** @var \DOMNode $attr */
                foreach ($child->attributes as $attr) {
                    [$ns, ] = explode(':', $attr->nodeName, 2);
                    if ($ns === 'style' || $ns === 'fo') {
                        $data[$attr->nodeName] = $attr->nodeValue;
                    }
                }
                $additionalData[$child->nodeName] = $data;
            }
        }

        if ($element->hasAttribute('style:data-style-name')) {
            $additionalData['style:data-style-name'] = $element->getAttribute('style:data-style-name');
        }

        if (!empty($additionalData)) {
            $styleData->setAdditionalData($additionalData);

            return $styleData;
        }

        return null;
    }

    /**
     * @param $styleName
     *
     * @return null|int
     */
    public function resolveStyleNameToIndex($styleName)
    {
        return isset($this->styleMap[$styleName]) ? $this->styleMap[$styleName] : null;
    }
}

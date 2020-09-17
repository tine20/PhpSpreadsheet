<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Ods;

use PhpOffice\PhpSpreadsheet\Exception;

class RowStyle extends AbstractStyle
{
    /**
     * see \PhpOffice\PhpSpreadsheet\Writer\Ods\Content::writeXfStyles.
     *
     * @param \DOMElement $node
     *
     * @throws Exception
     */
    public function readFromDom(\DOMElement $element)
    {
        $styleName = $element->getAttributeNS($this->styleNS, 'name');
        if (isset($this->styleMap[$styleName])) {
            throw new Exception('style name ' . $styleName . ' already in use');
        }

        /*
        $styleData = new Style();
        $allignment = $styleData->getAlignment();

        $children = $element->getElementsByTagNameNS($this->styleNS, 'table-row-properties');
        if ($children->length < 1) return;
        $element = $children->item(0);

        if (!empty($break = $element->getAttributeNS($this->fontNS, 'break-before'))) {
            $allignment->setBreakBefore($break);
        }

        if (!empty($height = $element->getAttributeNS($this->styleNS, 'row-height'))) {
            $allignment->setRowHeight($height);
        }

        if (!empty($optimalHeight = $element->getAttributeNS($this->styleNS, 'use-optimal-row-height'))) {
            $allignment->setUseOptimalRowHeight('true' === $optimalHeight ? true : false);
        }*/

        if (null !== ($styleData = $this->_readFromDom($element))) {
            $this->spreadsheet->addRowXf($styleData);
            $this->styleMap[$styleName] = $styleData->getIndex();
        }
    }
}

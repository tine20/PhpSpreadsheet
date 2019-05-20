<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Ods;

use PhpOffice\PhpSpreadsheet\Exception;

class ColumnStyle extends AbstractStyle
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

        $children = $element->getElementsByTagNameNS($this->styleNS, 'table-column-properties');
        if ($children->length < 1) return;
        /** @var \DOMElement $element *
        $element = $children->item(0);

        if (!empty($break = $element->getAttributeNS($this->fontNS, 'break-before'))) {
            $allignment->setBreakBefore($break);
        }

        if (!empty($colWidth = $element->getAttributeNS($this->styleNS, 'column-width'))) {
            $allignment->setColumnWidth($colWidth);
        }*/

        if (null !== ($styleData = $this->_readFromDom($element))) {
            $this->spreadsheet->addColumnXf($styleData);
            $this->styleMap[$styleName] = $styleData->getIndex();
        }
    }
}

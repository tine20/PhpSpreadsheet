<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Ods;

use PhpOffice\PhpSpreadsheet\Exception;

class TableStyle extends AbstractStyle
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

        if (null !== ($styleData = $this->_readFromDom($element))) {
            if (!empty($masterPage = $element->getAttributeNS($this->styleNS, 'master-page-name'))) {
                $styleData->getFill()->setFillType($masterPage);
            }
            $this->spreadsheet->addTableXf($styleData);
            $this->styleMap[$styleName] = $styleData->getIndex();
        }
    }
}

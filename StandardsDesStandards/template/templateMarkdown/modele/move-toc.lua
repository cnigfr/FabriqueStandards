function RawBlock(el)
  if el.format == "openxml" and el.text:match("<w:instrText[^>]*>TOC") then
    -- Supprime la TOC initiale
    return pandoc.Null
  end
end

function Div(el)
  if el.classes:includes("toc") then
    -- Réinsère un champ TOC Word (le vrai champ, comme le ferait Pandoc)
    local toc_xml = [[
<w:sdt>
  <w:sdtPr>
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:instrText xml:space="preserve">TOC \o "1-3" \h \z \u</w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>
]]
    return pandoc.RawBlock("openxml", toc_xml)
  end
end


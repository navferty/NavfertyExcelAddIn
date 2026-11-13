{
  "title": "Replace",
  "description": "Replace Cyrillic characters using different transformation methods.",
  "image": "https://navferty.net/images/Transliterate1.png"
}

[Home](/en/) › [Commands](/en/#commands) › Replace

# <img alt="" src="/images/icons/replace.png" class="icons" width="64" height="64" />Replace

Replace Cyrillic characters using different transformation methods.

<div class="subfeatures">

<div class="section">

<span id="withtransliteration"></span>

## <img alt="" src="/images/icons/transliterate.png" class="icons" width="64" height="64" /> With Transliteration

Complete transliteration of Cyrillic alphabet to English. For example: "Ж" → "Zh", "Щ" → "Shch". Based on ICAO Doc <a href="https://www.icao.int/sites/default/files/publications/DocSeries/9303_p3_cons_en.pdf" rel="noopener" target="_blank">9303</a>.

<div class="shots">

<img src="/images/Transliterate1.png" loading="lazy" width="419" height="206" alt="Transliteration - Before" /> <img src="/images/Transliterate2.png" loading="lazy" width="454" height="206" alt="Transliteration - After" />

</div>

</div>

<div class="section">

<span id="replacechars"></span>

## <img alt="" src="/images/icons/replaceChars.png" class="icons" width="64" height="64" /> Replace Chars

Replace visually similar Cyrillic letters with Latin letters: А→A, В→B, Е→E, К→K, М→M, Н→H, О→O, Р→P, С→C, Т→T, У→Y, Х→X.

<div class="shots">

<img src="/images/ReplaceChars1.png" loading="lazy" width="583" height="167" alt="Replace Chars - Before" /> <img src="/images/ReplaceChars2.png" loading="lazy" width="583" height="167" alt="Replace Chars - After" />

</div>

</div>

</div>

## How to use

Select the text cells and choose an option from **Replace**. **With Transliteration** converts Russian Cyrillic letters according to the displayed mapping; for example, `Жук` becomes `Zhuk`. **Replace Chars** changes only look-alike Cyrillic characters to Latin: `СОС` becomes `COC`. It does not convert Latin letters back to Cyrillic.

**Important:** Transliteration is not translation, and characters outside the supported mapping may not be retained. Review names against the required standard. The command writes values back to the selection and can replace formulas with their results; save a copy first.

The add-in is free and open source. [Installation instructions](/en/install/).

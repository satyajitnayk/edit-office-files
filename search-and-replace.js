import {DocxReader} from "./doc-reader.js";
import {
  findAllOccurrences,
  findRanges, getIndexOfSubstringEnd,
  getTextPieceFromTextRuns, LCSEndsWith2ndStrSubstring,
  LCSStartsWith2ndStrSubstring,
  updateTextPieceInTextRuns
} from "./utils.js";

export class SearchAndReplace {
  constructor(filePath, searchTexts, replacementTexts, destinationFilePath = null) {
    this.filePath = filePath;
    this.searchTexts = Array.isArray(searchTexts) ? searchTexts : [searchTexts];
    this.replacementTexts = Array.isArray(replacementTexts) ? replacementTexts : [replacementTexts];

    if (this.searchTexts.length !== this.replacementTexts.length) {
      throw new Error("searchTexts and replacementTexts arrays must have the same length.");
    }

    this.destinationFilePath = destinationFilePath || filePath;
  }

  async process() {
    const reader = new DocxReader(this.filePath);
    const xmlObject = await reader.readAsObject("word/document.xml");
    for (let i = 0; i < this.searchTexts.length; ++i) {
      this._searchAndReplace(xmlObject, this.searchTexts[i], this.replacementTexts[i]);
    }
    await reader.writeObject("word/document.xml", xmlObject);
    await reader.save(this.destinationFilePath);
  }

  _searchAndReplace(jsObject, searchText, replacementText) {
    const wpTags = jsObject['w:document']['w:body'][0]['w:p'];
    for (const [wpIndex, wpTag] of wpTags.entries()) {
      const ranges = [];
      let charIndex = 0;
      const wrTags = wpTag['w:r'] ?? [];
      if (wrTags.length === 0) {
        continue;
      }
      let completeText = '';
      wrTags.forEach((wrTag) => {
        const text = wrTag['w:t']?.[0]?.['_'] ?? wrTag['w:t']?.[0] ?? '';
        ranges.push([charIndex, charIndex + text.length - 1]);
        charIndex += text.length;
        completeText += text;
      });

      const wrIndexesToBeDeleted = [];

      // find all occurrence of search text inside completeText
      const occurrenceIndexes = findAllOccurrences(completeText, searchText);
      occurrenceIndexes.forEach((index) => {
        const startIndexOfSearchText = index;
        const endIndexOfSearchText = index + searchText.length - 1;
        const rangeIndexes = findRanges(
          ranges,
          startIndexOfSearchText,
          endIndexOfSearchText
        );

        let firstRangeIndexOfSearchText = rangeIndexes[0];
        let lastRangeIndexIndexOfSearchText = rangeIndexes[rangeIndexes.length - 1];

        // search text present in single text piece(<w:t>) tag
        if (rangeIndexes.length === 1) {
          const text = getTextPieceFromTextRuns(wrTags, firstRangeIndexOfSearchText);
          const updatedText = text.replace(searchText, replacementText);
          updateTextPieceInTextRuns(wrTags, firstRangeIndexOfSearchText, updatedText);
        } else {
          // search text expands over multiple text piece(w:t) tags
          // check if both part of searchText & text piece(w:t) does not starts with
          // same string in firstRangeIndexOfSearchText
          const firstTextPiece = getTextPieceFromTextRuns(wrTags, firstRangeIndexOfSearchText);
          if (!searchText.startsWith(firstTextPiece)) {
            // keep unmatched part of text piece(w:t) intact & replace remaining part with
            // replacement text & increment firstRangeIndexOfSearchText by 1
            const lcs = LCSStartsWith2ndStrSubstring(firstTextPiece, searchText);
            const indexOfLcs = firstTextPiece.indexOf(lcs);
            const updatedTextForFirstTextPiece = firstTextPiece.substring(0, indexOfLcs);
            updateTextPieceInTextRuns(wrTags, firstRangeIndexOfSearchText, updatedTextForFirstTextPiece);

            firstRangeIndexOfSearchText++;
          }

          updateTextPieceInTextRuns(wrTags, firstRangeIndexOfSearchText, replacementText);
          firstRangeIndexOfSearchText++;

          // check if both part of searchText & text piece(w:t) does not ends with
          // same string in lastRangeIndexOfSearchText
          const lastTextPiece = getTextPieceFromTextRuns(wrTags, lastRangeIndexIndexOfSearchText);
          if (!searchText.endsWith(lastTextPiece)) {
            // keep unmatched part of text piece(w:t) intact & decrement lastRangeIndexOfSearchText by 1
            const lcs = LCSEndsWith2ndStrSubstring(lastTextPiece, searchText);
            const indexOfLcsEnd = getIndexOfSubstringEnd(lastTextPiece, lcs);
            const updatedTextForFirstTextPiece = lastTextPiece.substring(indexOfLcsEnd + 1);
            updateTextPieceInTextRuns(wrTags, lastRangeIndexIndexOfSearchText, updatedTextForFirstTextPiece);

            lastRangeIndexIndexOfSearchText--;
          }

          // collect all intermediate range Indices of text piece(w:t) to be deleted later on
          if (firstRangeIndexOfSearchText <= lastRangeIndexIndexOfSearchText) {
            wrIndexesToBeDeleted.push([firstRangeIndexOfSearchText, lastRangeIndexIndexOfSearchText]);
          }
        }
      });

      // remove marked wrTags in reverse order -> this is done not to disturb original order
      wrIndexesToBeDeleted.reverse().forEach((range) => {
        wrTags.splice(range[0], range[1]);
      });

      jsObject['w:document']['w:body'][0]['w:p'][wpIndex]['w:r'] = wrTags;
    }
  }
}

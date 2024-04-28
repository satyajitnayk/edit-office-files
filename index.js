const fs = require('fs');
const xml2js = require('xml2js');

async function stringToObject(xmlString) {
  let error;
  let relation;
  return new Promise((resolve, reject) => {
    try {
      xml2js.parseString(xmlString, (err, rel) => {
        error = err;
        relation = rel;
      });
    } catch (e) {
      error = e;
    } finally {
      if (error) {
        reject(new Error(error.toString()));
      } else {
        resolve(relation);
      }
    }
  });
}

function objectToString(jsObject) {
  let builder = new xml2js.Builder();
  return builder.buildObject(jsObject);
}

const xml = fs.readFileSync('xmls/input.xml', 'utf-8');

const searchText = 'llo World How';
const replacementText = 'REPLACEMENT_STRING';

async function main() {
  const searchTextLength = searchText.length;
  const jsObject = await stringToObject(xml);

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
      const text = wrTag['w:t'][0]?.['_'] ?? wrTag['w:t'][0];
      ranges.push([charIndex, charIndex + text.length - 1]);
      charIndex += text.length;
      completeText += text;
    });

    const wrIndexesToBeDeleted = [];

    // find all occurrence of search text inside completeText
    const occurrenceIndexes = findAllOccurrences(completeText, searchText);
    occurrenceIndexes.forEach((index) => {
      const startIndexOfSearchText = index;
      const endIndexOfSearchText = index + searchTextLength - 1;
      const rangeIndexes = findRanges(
        ranges,
        startIndexOfSearchText,
        endIndexOfSearchText
      );

      let firstRangeIndexOfSearchText = rangeIndexes[0];
      let lastRangeIndexIndexOfSearchText = rangeIndexes[rangeIndexes.length - 1];

      // search text present in single text piece(<w:t>) tag
      if (rangeIndexes.length === 1) {
        const text = getTextPieceFromTextRuns(wrTags,firstRangeIndexOfSearchText);
        const updatedText = text.replace(searchText, replacementText);
        updateTextPieceInTextRuns(wrTags,firstRangeIndexOfSearchText, updatedText);
      } else {
        // search text expands over multiple text piece(w:t) tags
        // check if both part of searchText & text piece(w:t) does not starts with
        // same string in firstRangeIndexOfSearchText
        const firstTextPiece = getTextPieceFromTextRuns(wrTags,firstRangeIndexOfSearchText);
        if(!searchText.startsWith(firstTextPiece)) {
          // keep unmatched part of text piece(w:t) intact & replace remaining part with
          // replacement text & increment firstRangeIndexOfSearchText by 1
          const lcs = LCSStartsWith2ndStrSubstring(firstTextPiece, searchText);
          const indexOfLcs = firstTextPiece.indexOf(lcs);
          const updatedTextForFirstTextPiece = firstTextPiece.substring(0,indexOfLcs);
          updateTextPieceInTextRuns(wrTags,firstRangeIndexOfSearchText,updatedTextForFirstTextPiece);

          firstRangeIndexOfSearchText++;
        }


        updateTextPieceInTextRuns(wrTags,firstRangeIndexOfSearchText,replacementText);
        firstRangeIndexOfSearchText++;

        // check if both part of searchText & text piece(w:t) does not ends with
        // same string in lastRangeIndexOfSearchText
        const lastTextPiece = getTextPieceFromTextRuns(wrTags,lastRangeIndexIndexOfSearchText);
        if(!searchText.endsWith(lastTextPiece)) {
          // keep unmatched part of text piece(w:t) intact & decrement lastRangeIndexOfSearchText by 1
          const lcs = LCSEndsWith2ndStrSubstring(lastTextPiece, searchText);
          const indexOfLcsEnd = getIndexOfSubstringEnd(lastTextPiece, lcs);
          const updatedTextForFirstTextPiece = lastTextPiece.substring(indexOfLcsEnd+1);
          updateTextPieceInTextRuns(wrTags,lastRangeIndexIndexOfSearchText,updatedTextForFirstTextPiece);

          lastRangeIndexIndexOfSearchText--;
        }

        // collect all intermediate range Indices of text piece(w:t) to be deleted later on
        if(firstRangeIndexOfSearchText <= lastRangeIndexIndexOfSearchText) {
          wrIndexesToBeDeleted.push([firstRangeIndexOfSearchText,lastRangeIndexIndexOfSearchText]);
        }
      }
    });

    // remove marked wrTags in reverse order -> this is done not to disturb original order
    wrIndexesToBeDeleted.reverse().forEach((range) => {
      wrTags.splice(range[0], range[1]);
    });

    jsObject['w:document']['w:body'][0]['w:p'][wpIndex]['w:r'] = wrTags;
  }

  // convert object to xml string
  fs.writeFileSync('xmls/updated.xml', objectToString(jsObject));
}

function getTextPieceFromTextRuns(wrs, index) {
  if (wrs[index]['w:t'][0]['_']) {
    return wrs[index]['w:t'][0]['_'];
  } else if(wrs[index]['w:t'][0]){
    return wrs[index]['w:t'][0];
  } else {
    return '';
  }
}

function updateTextPieceInTextRuns(wrs, index, text) {
  if (wrs[index]['w:t'][0]['_']) {
     wrs[index]['w:t'][0]['_'] = text;
  } else{
     wrs[index]['w:t'][0] = text;
  }
}

/**
 * Finds and returns the indices of all ranges that overlap with the specified input range.
 *
 * This function performs a binary search to find the first range that overlaps with the
 * given input range and then collects all subsequent ranges that continue to overlap.
 *
 * @param {Array<Array<number>>} ranges - An array of arrays, where each sub-array represents a range with two elements [start, end].
 * @param {number} start - The start value of the input range.
 * @param {number} end - The end value of the input range.
 * @returns {Array<number>} An array of indices representing the ranges that overlap with the given input range.
 */
function findRanges(ranges, start, end) {
  let result = [];
  let low = 0;
  let high = ranges.length - 1;

  // Binary search
  while (low <= high) {
    let mid = Math.floor((low + high) / 2);
    if (ranges[mid][1] < start) {
      low = mid + 1;
    } else if (ranges[mid][0] > end) {
      high = mid - 1;
    } else {
      while (mid > 0 && ranges[mid - 1][1] >= start) {
        mid--;
      }

      while (mid < ranges.length && ranges[mid][0] <= end) {
        result.push(mid);
        mid++;
      }
      break;
    }
  }

  return result;
}

/**
 * Finds all indices of a substring in a given string.
 *
 * This function searches for all occurrences of a specified substring within a string
 * and returns an array of starting indices where the substring is found.
 *
 * @param {string} str - The string to search within.
 * @param {string} substr - The substring to search for.
 * @returns {Array<number>} An array of indices where the substring is found. If the substring is not found, an empty array is returned.
 */
function findAllOccurrences(str, substr) {
  let indices = [];
  let pos = str.indexOf(substr);

  while (pos !== -1) {
    indices.push(pos);
    pos = str.indexOf(substr, pos + 1); // Move the starting index forward to the next character after the last found index
  }

  return indices;
}

/**
 * longestCommonSubstringStartWith2ndStrSubString
 * @param str1
 * @param str2
 * @returns {string}
 */
function LCSStartsWith2ndStrSubstring(str1,str2) {
  let commonSubstring = "";
  for (let i = 0; i < str2.length; i++) {
    let substring = str2.substring(0, i + 1);
    if (str1.includes(substring)) {
      commonSubstring = substring;
    }
  }
  return commonSubstring;
}


/**
 * longestCommonSubstringEndsWith2ndStrSubString
 * @param str1
 * @param str2
 * @returns {string}
 */
function LCSEndsWith2ndStrSubstring(str1,str2) {
  let commonSubstring = "";
  for (let i = str2.length - 1; i >= 0; i--) {
    let substring = str2.substring(i);
    if (str1.startsWith(substring)) {
      commonSubstring = substring;
    }
  }
  return commonSubstring;
}

/**
 * find the index at which a substring ends in another string
 * @param str
 * @param substr
 * @returns {number}
 */
function getIndexOfSubstringEnd(str, substr) {
    return str.lastIndexOf(substr) + substr.length - 1;
}

main();

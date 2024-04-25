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

const searchText = 'Hello World';
const replacementText = 'REPLACEMENT_STRING';

async function main() {
  const searchTextLength = searchText.length;
  const jsObject = await stringToObject(xml);

  const wpTags = jsObject['w:document']['w:body'][0]['w:p'];
  for (const [wpIndex, wpTag] of wpTags.entries()) {
    const wrDeletionIndexRanges = [];
    const ranges = [];
    let charIndex = 0;
    const wrTags = wpTag['w:r'] ?? [];
    if (wrTags.length === 0) {
      continue;
    }
    const updatedWrTags = [...wrTags];
    let completeText = '';
    wrTags.forEach((wrTag) => {
      const text = wrTag['w:t'][0]?.['_'] ?? wrTag['w:t'][0];
      ranges.push([charIndex, charIndex + text.length - 1]);
      charIndex += text.length;
      completeText += text;
    });

    // find all occurence of search text inside completeText
    const occurenceIndexes = findAllOccurrences(completeText, searchText);
    occurenceIndexes.forEach((index) => {
      const startIndexOfSearchText = index;
      const endIndexOfSearchText = index + searchTextLength - 1;
      const rangeIndexes = findRanges(
        ranges,
        startIndexOfSearchText,
        endIndexOfSearchText
      );

      let firstIndex = rangeIndexes[0];
      let lastIndex = rangeIndexes[rangeIndexes.length - 1];

      // search text present in single text <w:t> tag
      if (rangeIndexes.length === 1) {
        if (updatedWrTags[firstIndex]['w:t'][0]['_']) {
          updatedWrTags[firstIndex]['w:t'][0]['_'] = updatedWrTags[firstIndex][
            'w:t'
          ][0]['_'].replace(searchText, replacementText);
        } else {
          updatedWrTags[firstIndex]['w:t'][0] = updatedWrTags[firstIndex][
            'w:t'
          ][0].replace(searchText, replacementText);
        }
      } else {
        // search text span across multiple <w:t> tags
        let wtTextDoesNotStartsWithSearchText = false;
        let wtTextDoesNotEndsWithSearchText = false;

        const rangeForStartOfSearchText = ranges[firstIndex];
        // if searchText does not start from start of wtText start, keep the unmatched text intact
        // wtText = "Hello "  search = "ello " keep "H" intact & replace "ello" with replacement text
        if (startIndexOfSearchText > rangeForStartOfSearchText[0]) {
          wtTextDoesNotStartsWithSearchText = true;
          if (updatedWrTags[firstIndex]['w:t'][0]['_']) {
            updatedWrTags[firstIndex]['w:t'][0]['_'] = updatedWrTags[
              firstIndex
            ]['w:t'][0]['_'].slice(
              rangeForStartOfSearchText[0],
              startIndexOfSearchText - rangeForStartOfSearchText[0]
            );
          } else {
            updatedWrTags[firstIndex]['w:t'][0] = updatedWrTags[firstIndex][
              'w:t'
            ][0].slice(
              rangeForStartOfSearchText[0],
              startIndexOfSearchText - rangeForStartOfSearchText[0]
            );
          }
        }

        const rangeForEndOfSearchText = ranges[lastIndex];
        // if searchText does not end with end of wtText end, keep the unmatched text intact
        // also wtText = "Hello "  search = "Hel" keep "lo " intact & replace "Hel" with replacement text
        if (endIndexOfSearchText < rangeForEndOfSearchText[1]) {
          wtTextDoesNotEndsWithSearchText = true;
          if (updatedWrTags[lastIndex]['w:t'][0]['_'] !== undefined) {
            updatedWrTags[lastIndex]['w:t'][0]['_'] = updatedWrTags[lastIndex][
              'w:t'
            ][0]['_'].slice(
              rangeForEndOfSearchText[1] - endIndexOfSearchText + 1
            );
          } else {
            updatedWrTags[lastIndex]['w:t'][0] = updatedWrTags[lastIndex][
              'w:t'
            ][0].slice(rangeForEndOfSearchText[1] - endIndexOfSearchText + 1);
          }
        }

        if (
          wtTextDoesNotStartsWithSearchText &&
          firstIndex < updatedWrTags.length - 1
        ) {
          firstIndex++;
        }

        if (wtTextDoesNotEndsWithSearchText && firstIndex > 0) {
          lastIndex--;
        }

        // replace firstIndex wt Text with replacement text
        if (updatedWrTags[firstIndex]['w:t'][0]['_'] !== undefined) {
          if (wtTextDoesNotStartsWithSearchText) {
            updatedWrTags[firstIndex]['w:t'][0]['_'] += replacementText;
          } else if (wtTextDoesNotEndsWithSearchText) {
            updatedWrTags[firstIndex]['w:t'][0]['_'] =
              replacementText + updatedWrTags[firstIndex]['w:t'][0]['_'];
          } else {
            updatedWrTags[firstIndex]['w:t'][0]['_'] = replacementText;
          }
        } else {
          if (wtTextDoesNotStartsWithSearchText) {
            updatedWrTags[firstIndex]['w:t'][0] += replacementText;
          } else if (wtTextDoesNotEndsWithSearchText) {
            updatedWrTags[firstIndex]['w:t'][0] =
              replacementText + updatedWrTags[firstIndex]['w:t'][0];
          } else {
            updatedWrTags[firstIndex]['w:t'][0] = replacementText;
          }
        }

        // remove remaining text runs
        wrDeletionIndexRanges.push([firstIndex + 1, lastIndex - firstIndex]);
      }
    });

    // remove marked wrTags in reverse order -> this is done to not to disturb original order
    wrDeletionIndexRanges.reverse().forEach((range) => {
      updatedWrTags.splice(range[0], range[1]);
    });

    jsObject['w:document']['w:body'][0]['w:p'][wpIndex]['w:r'] = updatedWrTags;
  }

  // convert object to xml string
  fs.writeFileSync('xmls/updated.xml', objectToString(jsObject));
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

main();

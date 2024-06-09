export function getTextPieceFromTextRuns(wrs, index) {
  if (wrs[index]['w:t'][0]['_']) {
    return wrs[index]['w:t'][0]['_'];
  } else if (wrs[index]['w:t'][0]) {
    return wrs[index]['w:t'][0];
  } else {
    return '';
  }
}

export function updateTextPieceInTextRuns(wrs, index, text) {
  if (wrs[index]['w:t'][0]['_']) {
    wrs[index]['w:t'][0]['_'] = text;
  } else {
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
export function findRanges(ranges, start, end) {
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
export function findAllOccurrences(str, substr) {
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
export function LCSStartsWith2ndStrSubstring(str1, str2) {
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
export function LCSEndsWith2ndStrSubstring(str1, str2) {
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
export function getIndexOfSubstringEnd(str, substr) {
  return str.lastIndexOf(substr) + substr.length - 1;
}

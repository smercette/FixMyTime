/**
 * Test the improved date parsing functionality (JavaScript version)
 */

// Updated parseDate function matching the one in taskpane.ts
function parseDate(dateValue) {
  if (!dateValue) return null;
  
  // If it's already a Date object, return it
  if (dateValue instanceof Date) {
    return isNaN(dateValue.getTime()) ? null : dateValue;
  }
  
  // Convert to string and trim
  const dateStr = dateValue.toString().trim();
  if (!dateStr) return null;
  
  // Try parsing Excel serial number (number of days since 1900-01-01)
  const numValue = Number(dateStr);
  if (!isNaN(numValue) && numValue > 0 && numValue < 100000 && dateStr.match(/^\d+$/)) {
    // Excel's epoch is January 1, 1900 (serial number 1)
    const epoch = new Date(1900, 0, 1);
    
    // Excel has a bug where it thinks 1900 is a leap year
    // If the serial is after February 28, 1900 (serial 59), we need to subtract 1 day
    let adjustedSerial = numValue;
    if (numValue > 59) {
      adjustedSerial = numValue - 1;
    }
    
    // Convert to date (subtract 1 because Excel serial 1 = Jan 1, 1900)
    return new Date(epoch.getTime() + (adjustedSerial - 1) * 24 * 60 * 60 * 1000);
  }
  
  // Define date format patterns with explicit parsing
  const patterns = [
    // ISO formats (unambiguous)
    {
      regex: /^(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+\d{1,2}:\d{1,2})?$/,
      parse: (match) => new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]))
    },
    {
      regex: /^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s+\d{1,2}:\d{1,2})?$/,
      parse: (match) => new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]))
    },
    
    // European format with dots (DD.MM.YYYY - unambiguous due to separator)
    {
      regex: /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/,
      parse: (match) => new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]))
    },
    
    // Ambiguous slash/dash formats
    {
      regex: /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+\d{1,2}:\d{1,2}.*)?$/,
      parse: (match) => {
        const [, first, second, year] = match;
        const f = parseInt(first);
        const s = parseInt(second);
        const y = parseInt(year);
        
        // Try to determine format based on values
        if (f > 12 && s <= 12) {
          // First number > 12, must be day (DD/MM/YYYY)
          return new Date(y, s - 1, f);
        } else if (s > 12 && f <= 12) {
          // Second number > 12, must be day (MM/DD/YYYY)
          return new Date(y, f - 1, s);
        } else {
          // Default to MM/DD/YYYY (US format)
          return new Date(y, f - 1, s);
        }
      }
    }
  ];
  
  // Try each pattern
  for (const pattern of patterns) {
    const match = dateStr.match(pattern.regex);
    if (match) {
      try {
        const result = pattern.parse(match);
        if (result && !isNaN(result.getTime())) {
          return result;
        }
      } catch (e) {
        // Continue to next pattern
      }
    }
  }
  
  // Try built-in Date parsing as last resort (good for natural language dates)
  let parsed = new Date(dateStr.replace(/(\d+)(st|nd|rd|th)/g, '$1')); // Remove ordinals
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  
  // Final fallback - try standard Date parsing
  parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  
  return null;
}

// Test the new datesWithinTolerance function
function datesWithinTolerance(date1, date2, toleranceDays) {
  try {
    const d1 = parseDate(date1);
    const d2 = parseDate(date2);

    if (!d1 || !d2) {
      console.log(`Date parsing failed: date1="${date1}" parsed to ${d1}, date2="${date2}" parsed to ${d2}`);
      return false;
    }

    const diffMs = Math.abs(d1.getTime() - d2.getTime());
    const diffDays = diffMs / (1000 * 60 * 60 * 24);

    return diffDays <= toleranceDays;
  } catch (error) {
    console.error(`Error comparing dates: ${error}`);
    return false;
  }
}

// Format date for display
function formatDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

// Run tests
console.log('=== Testing Improved Date Parsing v2 ===\n');

// Test various date formats
const testDates = [
  // ISO formats
  { input: '2024-01-06', expected: '2024-01-06', desc: 'ISO format YYYY-MM-DD' },
  { input: '2024/01/06', expected: '2024-01-06', desc: 'ISO format YYYY/MM/DD' },
  
  // US formats
  { input: '01/06/2024', expected: '2024-01-06', desc: 'US format MM/DD/YYYY' },
  { input: '1/6/2024', expected: '2024-01-06', desc: 'US format M/D/YYYY' },
  { input: '01-06-2024', expected: '2024-01-06', desc: 'US format MM-DD-YYYY' },
  
  // European formats
  { input: '06.01.2024', expected: '2024-01-06', desc: 'European format DD.MM.YYYY' },
  { input: '6.1.2024', expected: '2024-01-06', desc: 'European format D.M.YYYY' },
  
  // With timestamps
  { input: '2024-01-06 14:30', expected: '2024-01-06', desc: 'ISO with time' },
  { input: '01/06/2024 2:30 PM', expected: '2024-01-06', desc: 'US with time' },
  
  // Excel serial numbers (45297 = January 6, 2024)
  { input: '45297', expected: '2024-01-06', desc: 'Excel serial number' },
  { input: 45297, expected: '2024-01-06', desc: 'Excel serial number (numeric)' },
  
  // Natural language formats
  { input: '6 January 2024', expected: '2024-01-06', desc: 'Natural language' },
  { input: 'Jan 6, 2024', expected: '2024-01-06', desc: 'Short month name' },
  { input: '6th January 2024', expected: '2024-01-06', desc: 'With ordinal' },
  
  // Invalid formats
  { input: 'invalid', expected: null, desc: 'Invalid date string' },
  { input: '', expected: null, desc: 'Empty string' },
  { input: null, expected: null, desc: 'Null value' },
];

console.log('Testing individual date parsing:');
let passedTests = 0;
testDates.forEach(test => {
  const parsed = parseDate(test.input);
  const result = parsed ? formatDate(parsed) : null;
  const status = result === test.expected ? '✓' : '✗';
  if (result === test.expected) passedTests++;
  console.log(`${status} ${test.desc}: "${test.input}" → ${result}`);
  if (result !== test.expected) {
    console.log(`  Expected: ${test.expected}`);
  }
});

console.log(`\nIndividual parsing tests: ${passedTests}/${testDates.length} passed\n`);

// Test date comparison with tolerance
console.log('=== Testing Date Comparison with Tolerance ===\n');

const comparisonTests = [
  { d1: '2024-01-06', d2: '2024-01-06', tol: 0, expected: true, desc: 'Same date (ISO)' },
  { d1: '01/06/2024', d2: '2024-01-06', tol: 0, expected: true, desc: 'Different formats, same date' },
  { d1: '06.01.2024', d2: '1/6/2024', tol: 0, expected: true, desc: 'EU vs US format, same date' },
  { d1: '45297', d2: '2024-01-06', tol: 0, expected: true, desc: 'Excel serial vs ISO' },
  { d1: '2024-01-06', d2: '2024-01-07', tol: 1, expected: true, desc: 'Next day, 1 day tolerance' },
  { d1: '2024-01-06', d2: '2024-01-08', tol: 1, expected: false, desc: '2 days apart, 1 day tolerance' },
  { d1: 'Jan 6, 2024', d2: '6 January 2024', tol: 0, expected: true, desc: 'Natural language formats' },
];

let passedComparisons = 0;
comparisonTests.forEach(test => {
  const result = datesWithinTolerance(test.d1, test.d2, test.tol);
  const status = result === test.expected ? '✓' : '✗';
  if (result === test.expected) passedComparisons++;
  console.log(`${status} ${test.desc}: ${result}`);
  if (result !== test.expected) {
    console.log(`  Expected: ${test.expected}`);
  }
});

console.log(`\nComparison tests: ${passedComparisons}/${comparisonTests.length} passed\n`);

console.log('=== Date Parsing Test v2 Complete ===');
console.log(`Overall: ${passedTests + passedComparisons}/${testDates.length + comparisonTests.length} tests passed`);
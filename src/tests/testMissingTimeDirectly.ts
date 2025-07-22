/**
 * Direct test runner for Missing Time functionality
 * This can be run directly without a test framework
 */

// Test the date tolerance function
function testDatesWithinTolerance() {
  console.log('\n=== Testing datesWithinTolerance ===');
  
  const datesWithinTolerance = (date1: any, date2: any, toleranceDays: number): boolean => {
    try {
      const d1 = new Date(date1);
      const d2 = new Date(date2);
      
      if (isNaN(d1.getTime()) || isNaN(d2.getTime())) {
        return false;
      }
      
      const diffMs = Math.abs(d1.getTime() - d2.getTime());
      const diffDays = diffMs / (1000 * 60 * 60 * 24);
      
      return diffDays <= toleranceDays;
    } catch {
      return false;
    }
  };
  
  // Test cases
  const tests = [
    { d1: '2024-01-06', d2: '2024-01-06', tol: 0, expected: true, desc: 'Same date, 0 tolerance' },
    { d1: '2024-01-06', d2: '2024-01-07', tol: 1, expected: true, desc: 'Next day, 1 day tolerance' },
    { d1: '2024-01-06', d2: '2024-01-08', tol: 1, expected: false, desc: '2 days apart, 1 day tolerance' },
    { d1: '01/06/2024', d2: '06/01/2024', tol: 0, expected: true, desc: 'Different date formats (MM/DD vs DD/MM)' },
    { d1: 'invalid', d2: '2024-01-06', tol: 0, expected: false, desc: 'Invalid date' },
    { d1: null, d2: '2024-01-06', tol: 0, expected: false, desc: 'Null date' }
  ];
  
  tests.forEach(test => {
    const result = datesWithinTolerance(test.d1, test.d2, test.tol);
    const status = result === test.expected ? '✓' : '✗';
    console.log(`${status} ${test.desc}: ${result}`);
    if (result !== test.expected) {
      console.log(`  Expected: ${test.expected}, Got: ${result}`);
    }
  });
}

// Test name matching
function testNameMatching() {
  console.log('\n=== Testing Name Matching ===');
  
  const testMatch = (narrative: string, feeEarnerName: string): boolean => {
    const firstName = feeEarnerName.split(" ")[0].toLowerCase();
    const lastName = feeEarnerName.split(" ").slice(1).join(" ").toLowerCase();
    const fullName = feeEarnerName.toLowerCase();
    
    const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    
    const firstNameRegex = new RegExp(`\\b${escapeRegex(firstName)}\\b`, "i");
    const lastNameRegex = lastName ? new RegExp(`\\b${escapeRegex(lastName)}\\b`, "i") : null;
    const fullNameRegex = new RegExp(`\\b${escapeRegex(fullName)}\\b`, "i");
    
    const narrativeLower = narrative.toLowerCase();
    
    return fullNameRegex.test(narrativeLower) || 
           (firstNameRegex.test(narrativeLower) && 
            lastNameRegex && lastNameRegex.test(narrativeLower));
  };
  
  const tests = [
    { narrative: 'Call with Callum Reyes re disclosure', name: 'Callum Reyes', expected: true },
    { narrative: 'Meeting with Sophie, Callum, and Theo', name: 'Callum Reyes', expected: false },
    { narrative: 'Call re disclosure', name: 'Callum Reyes', expected: false },
    { narrative: 'CALL WITH CALLUM REYES', name: 'Callum Reyes', expected: true },
    { narrative: 'Callum mentioned something', name: 'Callum Reyes', expected: false },
    { narrative: 'Discussion with Reyes about case', name: 'Callum Reyes', expected: false }
  ];
  
  tests.forEach(test => {
    const result = testMatch(test.narrative, test.name);
    const status = result === test.expected ? '✓' : '✗';
    console.log(`${status} "${test.narrative}" matches "${test.name}": ${result}`);
    if (result !== test.expected) {
      console.log(`  Expected: ${test.expected}, Got: ${result}`);
    }
  });
}

// Test keyword matching
function testKeywordMatching() {
  console.log('\n=== Testing Keyword Matching ===');
  
  const testKeyword = (narrative: string, keyword: string): boolean => {
    const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const keywordRegex = new RegExp(`\\b${escapeRegex(keyword.toLowerCase())}\\b`, "i");
    return keywordRegex.test(narrative.toLowerCase());
  };
  
  const tests = [
    { narrative: 'Call with client', keyword: 'call', expected: true },
    { narrative: 'Team meeting re strategy', keyword: 'meeting', expected: true },
    { narrative: 'Recall previous discussion', keyword: 'call', expected: false },
    { narrative: 'Call/meeting with team', keyword: 'call', expected: true },
    { narrative: 'CONFERENCE with team', keyword: 'conference', expected: true },
    { narrative: 'Discussing phone options', keyword: 'phone', expected: true },
    { narrative: 'telephone conference', keyword: 'telephone', expected: true }
  ];
  
  tests.forEach(test => {
    const result = testKeyword(test.narrative, test.keyword);
    const status = result === test.expected ? '✓' : '✗';
    console.log(`${status} "${test.narrative}" contains keyword "${test.keyword}": ${result}`);
    if (result !== test.expected) {
      console.log(`  Expected: ${test.expected}, Got: ${result}`);
    }
  });
}

// Test the complete Missing Time logic
function testMissingTimeScenarios() {
  console.log('\n=== Testing Missing Time Scenarios ===');
  
  interface TimeEntry {
    date: string;
    feeEarner: string;
    narrative: string;
  }
  
  const mockFeeEarners = [
    { name: "Sophie Whitmore" },
    { name: "Callum Reyes" },
    { name: "Theo Johnson" },
    { name: "Chloe Anders" }
  ];
  
  const meetingKeywords = ["meeting", "call", "conference", "discussion", "telephone", "phone"];
  
  // Simulate the missing time detection logic
  const findMissing = (entries: TimeEntry[]) => {
    const results: string[] = [];
    
    entries.forEach((entry, idx) => {
      // Check for meeting keyword
      const narrative = entry.narrative.toLowerCase();
      const hasMeetingKeyword = meetingKeywords.some(kw => {
        const regex = new RegExp(`\\b${kw}\\b`, 'i');
        return regex.test(narrative);
      });
      
      if (!hasMeetingKeyword) return;
      
      // Find mentioned fee earners
      const mentioned = mockFeeEarners.filter(fe => {
        if (fe.name === entry.feeEarner) return false; // Skip self
        
        const fullName = fe.name.toLowerCase();
        const firstName = fe.name.split(' ')[0].toLowerCase();
        const lastName = fe.name.split(' ').slice(1).join(' ').toLowerCase();
        
        const fullNameRegex = new RegExp(`\\b${fullName}\\b`, 'i');
        const firstNameRegex = new RegExp(`\\b${firstName}\\b`, 'i');
        const lastNameRegex = lastName ? new RegExp(`\\b${lastName}\\b`, 'i') : null;
        
        return fullNameRegex.test(narrative) || 
               (firstNameRegex.test(narrative) && lastNameRegex && lastNameRegex.test(narrative));
      });
      
      // Check for reciprocal entries
      mentioned.forEach(mentionedFE => {
        const hasReciprocal = entries.some(other => {
          if (other.feeEarner !== mentionedFE.name) return false;
          if (other.date !== entry.date) return false;
          
          const otherNarrative = other.narrative.toLowerCase();
          
          // Check if mentions original fee earner
          const origFirstName = entry.feeEarner.split(' ')[0].toLowerCase();
          const origFullName = entry.feeEarner.toLowerCase();
          
          const mentionsOrig = new RegExp(`\\b${origFirstName}\\b`, 'i').test(otherNarrative) ||
                              new RegExp(`\\b${origFullName}\\b`, 'i').test(otherNarrative);
          
          // Check if has meeting keyword
          const hasMeetingKw = meetingKeywords.some(kw => {
            return new RegExp(`\\b${kw}\\b`, 'i').test(otherNarrative);
          });
          
          return mentionsOrig && hasMeetingKw;
        });
        
        if (!hasReciprocal) {
          results.push(`Row ${idx + 1} (${entry.feeEarner}): Missing Time - ${mentionedFE.name} should have entry for ${entry.date}`);
        }
      });
    });
    
    return results;
  };
  
  // Test scenarios
  console.log('\nScenario 1: Basic missing reciprocal');
  const scenario1: TimeEntry[] = [
    { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Call with Callum Reyes re disclosure' },
    { date: '2024-01-06', feeEarner: 'Callum Reyes', narrative: 'Review partnership agreement' }
  ];
  const result1 = findMissing(scenario1);
  console.log(result1.length === 1 ? '✓' : '✗', 'Should find 1 missing entry');
  result1.forEach(r => console.log('  ', r));
  
  console.log('\nScenario 2: Valid reciprocal exists');
  const scenario2: TimeEntry[] = [
    { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Call with Callum Reyes re disclosure' },
    { date: '2024-01-06', feeEarner: 'Callum Reyes', narrative: 'Call with Sophie about disclosure' }
  ];
  const result2 = findMissing(scenario2);
  console.log(result2.length === 0 ? '✓' : '✗', 'Should find 0 missing entries');
  result2.forEach(r => console.log('  ', r));
  
  console.log('\nScenario 3: Multiple participants, one missing');
  const scenario3: TimeEntry[] = [
    { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Meeting with Callum Reyes and Theo Johnson' },
    { date: '2024-01-06', feeEarner: 'Callum Reyes', narrative: 'Meeting with Sophie Whitmore and Theo Johnson' }
  ];
  const result3 = findMissing(scenario3);
  console.log(result3.length === 2 ? '✓' : '✗', 'Should find 2 missing entries (both for Theo)');
  result3.forEach(r => console.log('  ', r));
  
  console.log('\nScenario 4: No meeting keyword');
  const scenario4: TimeEntry[] = [
    { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Review documents with Callum Reyes' }
  ];
  const result4 = findMissing(scenario4);
  console.log(result4.length === 0 ? '✓' : '✗', 'Should find 0 missing entries (no meeting keyword)');
  result4.forEach(r => console.log('  ', r));
}

// Run all tests
console.log('=== Running Missing Time Rule Tests ===');
testDatesWithinTolerance();
testNameMatching();
testKeywordMatching();
testMissingTimeScenarios();
console.log('\n=== Tests Complete ===');
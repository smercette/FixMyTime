/**
 * Comprehensive tests for Missing Time Rule functionality
 */

// Import the functions we need to test
import { datesWithinTolerance, addNoteToRow, findNotesColumn } from '../taskpane/taskpane';

// Mock data for testing
const mockFeeEarners = [
  { name: "Sophie Whitmore" },
  { name: "Callum Reyes" },
  { name: "Theo Johnson" },
  { name: "Chloe Anders" },
  { name: "Nathan Cole" }
];

const mockMeetingKeywords = ["meeting", "call", "conference", "discussion", "telephone", "phone"];

// Test cases for datesWithinTolerance function
describe('datesWithinTolerance', () => {
  test('should return true for same dates', () => {
    const result = datesWithinTolerance('2024-01-06', '2024-01-06', 0);
    expect(result).toBe(true);
  });

  test('should return true for dates within tolerance', () => {
    const result = datesWithinTolerance('2024-01-06', '2024-01-07', 1);
    expect(result).toBe(true);
  });

  test('should return false for dates outside tolerance', () => {
    const result = datesWithinTolerance('2024-01-06', '2024-01-08', 1);
    expect(result).toBe(false);
  });

  test('should handle different date formats', () => {
    const result = datesWithinTolerance('01/06/2024', '2024-01-06', 0);
    expect(result).toBe(true);
  });

  test('should return false for invalid dates', () => {
    const result = datesWithinTolerance('invalid-date', '2024-01-06', 0);
    expect(result).toBe(false);
  });

  test('should handle null/undefined dates', () => {
    const result = datesWithinTolerance(null, '2024-01-06', 0);
    expect(result).toBe(false);
  });
});

// Test cases for addNoteToRow function
describe('addNoteToRow', () => {
  test('should add note to empty cell', () => {
    const result = addNoteToRow('', 'Missing Time: Test');
    expect(result).toBe('Missing Time: Test');
  });

  test('should append note to existing notes', () => {
    const result = addNoteToRow('Name Standardised', 'Missing Time: Test');
    expect(result).toBe('Name Standardised, Missing Time: Test');
  });

  test('should not duplicate existing notes', () => {
    const existing = 'Name Standardised, Missing Time: Test';
    const result = addNoteToRow(existing, 'Missing Time: Test');
    expect(result).toBe(existing);
  });

  test('should handle whitespace properly', () => {
    const result = addNoteToRow('  ', 'Missing Time: Test');
    expect(result).toBe('Missing Time: Test');
  });
});

// Test cases for name matching in narratives
describe('Fee Earner Name Matching', () => {
  const testNameMatching = (narrative: string, feeEarnerName: string): boolean => {
    const firstName = feeEarnerName.split(" ")[0].toLowerCase();
    const lastName = feeEarnerName.split(" ").slice(1).join(" ").toLowerCase();
    const fullName = feeEarnerName.toLowerCase();
    
    const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    
    const firstNameRegex = new RegExp(`\\b${escapeRegex(firstName)}\\b`, "i");
    const lastNameRegex = lastName ? new RegExp(`\\b${escapeRegex(lastName)}\\b`, "i") : null;
    const fullNameRegex = new RegExp(`\\b${escapeRegex(fullName)}\\b`, "i");
    
    return fullNameRegex.test(narrative.toLowerCase()) || 
           (firstNameRegex.test(narrative.toLowerCase()) && 
            lastNameRegex && lastNameRegex.test(narrative.toLowerCase()));
  };

  test('should match full name', () => {
    const result = testNameMatching('Call with Callum Reyes re disclosure', 'Callum Reyes');
    expect(result).toBe(true);
  });

  test('should match name with punctuation', () => {
    const result = testNameMatching('Meeting with Sophie, Callum, and Theo', 'Callum Reyes');
    expect(result).toBe(false); // Only first name, not full match
  });

  test('should not match partial names', () => {
    const result = testNameMatching('Call re disclosure', 'Callum Reyes');
    expect(result).toBe(false);
  });

  test('should handle names with special characters', () => {
    const result = testNameMatching("Call with O'Brien re matter", "O'Brien");
    expect(result).toBe(true);
  });

  test('should be case insensitive', () => {
    const result = testNameMatching('CALL WITH CALLUM REYES', 'Callum Reyes');
    expect(result).toBe(true);
  });
});

// Test cases for meeting keyword matching
describe('Meeting Keyword Matching', () => {
  const testKeywordMatching = (narrative: string, keyword: string): boolean => {
    const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const keywordRegex = new RegExp(`\\b${escapeRegex(keyword.toLowerCase())}\\b`, "i");
    return keywordRegex.test(narrative.toLowerCase());
  };

  test('should match "call" keyword', () => {
    const result = testKeywordMatching('Call with client', 'call');
    expect(result).toBe(true);
  });

  test('should match "meeting" keyword', () => {
    const result = testKeywordMatching('Team meeting re strategy', 'meeting');
    expect(result).toBe(true);
  });

  test('should not match partial words', () => {
    const result = testKeywordMatching('Recall previous discussion', 'call');
    expect(result).toBe(false);
  });

  test('should match with punctuation', () => {
    const result = testKeywordMatching('Call/meeting with team', 'call');
    expect(result).toBe(true);
  });

  test('should be case insensitive', () => {
    const result = testKeywordMatching('CONFERENCE with team', 'conference');
    expect(result).toBe(true);
  });
});

// Integration test scenarios
describe('Missing Time Rule Scenarios', () => {
  // Simulated entry data structure
  interface TimeEntry {
    date: string;
    feeEarner: string;
    narrative: string;
    rowIndex: number;
  }

  const findMissingEntries = (entries: TimeEntry[], feeEarners: any[], meetingKeywords: string[]) => {
    const missingEntries: any[] = [];
    
    entries.forEach((entry, index) => {
      // Check if contains meeting keyword
      const narrative = entry.narrative.toLowerCase();
      const containsMeetingKeyword = meetingKeywords.some(keyword => {
        const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const keywordRegex = new RegExp(`\\b${escapeRegex(keyword.toLowerCase())}\\b`, "i");
        return keywordRegex.test(narrative);
      });

      if (containsMeetingKeyword) {
        // Find mentioned fee earners
        const mentionedFeeEarners = feeEarners.filter(feeEarner => {
          const firstName = feeEarner.name.split(" ")[0].toLowerCase();
          const lastName = feeEarner.name.split(" ").slice(1).join(" ").toLowerCase();
          const fullName = feeEarner.name.toLowerCase();
          
          const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          
          const firstNameRegex = new RegExp(`\\b${escapeRegex(firstName)}\\b`, "i");
          const lastNameRegex = lastName ? new RegExp(`\\b${escapeRegex(lastName)}\\b`, "i") : null;
          const fullNameRegex = new RegExp(`\\b${escapeRegex(fullName)}\\b`, "i");
          
          const ismentioned = fullNameRegex.test(narrative) || 
                             (firstNameRegex.test(narrative) && lastNameRegex && lastNameRegex.test(narrative));
          
          return ismentioned && feeEarner.name !== entry.feeEarner;
        });

        // Check for reciprocal entries
        mentionedFeeEarners.forEach(mentionedFeeEarner => {
          const hasReciprocalEntry = entries.some(otherEntry => {
            if (otherEntry.feeEarner !== mentionedFeeEarner.name) return false;
            if (!datesWithinTolerance(entry.date, otherEntry.date, 0)) return false;
            
            const otherNarrative = otherEntry.narrative.toLowerCase();
            const originalFirstName = entry.feeEarner.split(" ")[0].toLowerCase();
            const originalFullName = entry.feeEarner.toLowerCase();
            
            const escapeRegex = (str: string) => str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            
            const firstNameRegex = new RegExp(`\\b${escapeRegex(originalFirstName)}\\b`, "i");
            const fullNameRegex = new RegExp(`\\b${escapeRegex(originalFullName)}\\b`, "i");
            const mentionsOriginalFeeEarner = firstNameRegex.test(otherNarrative) || fullNameRegex.test(otherNarrative);
            
            const containsMeetingKeyword = meetingKeywords.some(keyword => {
              const keywordRegex = new RegExp(`\\b${escapeRegex(keyword.toLowerCase())}\\b`, "i");
              return keywordRegex.test(otherNarrative);
            });
            
            return mentionsOriginalFeeEarner && containsMeetingKeyword;
          });

          if (!hasReciprocalEntry) {
            missingEntries.push({
              originalEntry: entry,
              missingFeeEarner: mentionedFeeEarner,
              date: entry.date,
              narrative: entry.narrative
            });
          }
        });
      }
    });

    return missingEntries;
  };

  test('Scenario 1: Basic missing reciprocal entry', () => {
    const entries: TimeEntry[] = [
      { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Call with Callum Reyes re disclosure', rowIndex: 1 },
      { date: '2024-01-06', feeEarner: 'Callum Reyes', narrative: 'Review partnership agreement', rowIndex: 2 }
    ];

    const missing = findMissingEntries(entries, mockFeeEarners, mockMeetingKeywords);
    
    expect(missing.length).toBe(1);
    expect(missing[0].originalEntry.feeEarner).toBe('Sophie Whitmore');
    expect(missing[0].missingFeeEarner.name).toBe('Callum Reyes');
  });

  test('Scenario 2: Valid reciprocal entry exists', () => {
    const entries: TimeEntry[] = [
      { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Call with Callum Reyes re disclosure', rowIndex: 1 },
      { date: '2024-01-06', feeEarner: 'Callum Reyes', narrative: 'Call with Sophie about disclosure', rowIndex: 2 }
    ];

    const missing = findMissingEntries(entries, mockFeeEarners, mockMeetingKeywords);
    
    expect(missing.length).toBe(0);
  });

  test('Scenario 3: Multiple participants with one missing', () => {
    const entries: TimeEntry[] = [
      { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Meeting with Callum Reyes and Theo Johnson re strategy', rowIndex: 1 },
      { date: '2024-01-06', feeEarner: 'Callum Reyes', narrative: 'Meeting with Sophie Whitmore and Theo Johnson', rowIndex: 2 }
    ];

    const missing = findMissingEntries(entries, mockFeeEarners, mockMeetingKeywords);
    
    expect(missing.length).toBe(2); // Both Sophie and Callum should have notes about Theo
    expect(missing.filter(m => m.missingFeeEarner.name === 'Theo Johnson').length).toBe(2);
  });

  test('Scenario 4: Date mismatch', () => {
    const entries: TimeEntry[] = [
      { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Call with Callum Reyes', rowIndex: 1 },
      { date: '2024-01-07', feeEarner: 'Callum Reyes', narrative: 'Call with Sophie', rowIndex: 2 }
    ];

    const missing = findMissingEntries(entries, mockFeeEarners, mockMeetingKeywords);
    
    expect(missing.length).toBe(2); // Both should be flagged as missing
  });

  test('Scenario 5: No meeting keyword in narrative', () => {
    const entries: TimeEntry[] = [
      { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Review documents with Callum Reyes', rowIndex: 1 }
    ];

    const missing = findMissingEntries(entries, mockFeeEarners, mockMeetingKeywords);
    
    expect(missing.length).toBe(0); // No meeting keyword, so not processed
  });

  test('Scenario 6: Case sensitivity handling', () => {
    const entries: TimeEntry[] = [
      { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'CALL WITH CALLUM REYES', rowIndex: 1 },
      { date: '2024-01-06', feeEarner: 'Callum Reyes', narrative: 'call with sophie whitmore', rowIndex: 2 }
    ];

    const missing = findMissingEntries(entries, mockFeeEarners, mockMeetingKeywords);
    
    expect(missing.length).toBe(0); // Should match despite case differences
  });

  test('Scenario 7: Fee earner not in list', () => {
    const entries: TimeEntry[] = [
      { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Call with Unknown Person', rowIndex: 1 }
    ];

    const missing = findMissingEntries(entries, mockFeeEarners, mockMeetingKeywords);
    
    expect(missing.length).toBe(0); // Unknown Person not in fee earners list
  });

  test('Scenario 8: Self-reference should be ignored', () => {
    const entries: TimeEntry[] = [
      { date: '2024-01-06', feeEarner: 'Sophie Whitmore', narrative: 'Call with Sophie Whitmore', rowIndex: 1 }
    ];

    const missing = findMissingEntries(entries, mockFeeEarners, mockMeetingKeywords);
    
    expect(missing.length).toBe(0); // Should not flag self-references
  });
});

// Test for Notes column finding
describe('findNotesColumn', () => {
  test('should find Notes column', () => {
    const headers = ['Date', 'Fee Earner', 'Narrative', 'Notes'];
    const result = findNotesColumn(headers);
    expect(result).toBe(3);
  });

  test('should find column with different case', () => {
    const headers = ['Date', 'Fee Earner', 'Narrative', 'NOTES'];
    const result = findNotesColumn(headers);
    expect(result).toBe(3);
  });

  test('should find "Rules Applied" column', () => {
    const headers = ['Date', 'Fee Earner', 'Narrative', 'Rules Applied'];
    const result = findNotesColumn(headers);
    expect(result).toBe(3);
  });

  test('should return -1 if not found', () => {
    const headers = ['Date', 'Fee Earner', 'Narrative'];
    const result = findNotesColumn(headers);
    expect(result).toBe(-1);
  });
});

// Run all tests
if (require.main === module) {
  console.log('Running Missing Time Rule Tests...');
  
  // Since we don't have a test runner set up, we'll do a simple implementation
  const runTests = () => {
    let passed = 0;
    let failed = 0;
    
    // You would normally use Jest or another test framework
    console.log('✓ All test scenarios defined');
    console.log('To run these tests properly, install Jest and run: npm test');
  };
  
  runTests();
}

export { testNameMatching, testKeywordMatching, findMissingEntries };
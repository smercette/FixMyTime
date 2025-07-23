/**
 * Test for Missing Time rule name swapping functionality
 */

// Test the name swapping logic
function testNameSwapping() {
  console.log('\n=== Testing Name Swapping in Missing Time Rule ===');
  
  // Mock fee earners list with full names
  const feeEarners = [
    { name: "Sophie Smith", role: "Senior Associate" },
    { name: "Callum Reyes", role: "Partner" },
    { name: "Theo Johnson", role: "Associate" },
    { name: "William Chen", role: "Senior Associate" }
  ];
  
  // Function to simulate the name swapping logic
  const swapNamesInNarrative = (
    originalNarrative: string, 
    missingFeeEarnerName: string,
    originalFeeEarnerFromEntry: string,
    feeEarnersList: Array<{name: string, role: string}>
  ): string => {
    // Find the full name of the original fee earner from the fee earners list
    let originalFeeEarnerFullName = originalFeeEarnerFromEntry;
    const originalFirstName = originalFeeEarnerFromEntry.split(" ")[0].toLowerCase();
    
    // Look for a match in the fee earners list to get the proper full name
    const matchingFeeEarner = feeEarnersList.find((fe) => {
      const feFirstName = fe.name.split(" ")[0].toLowerCase();
      const feFullNameLower = fe.name.toLowerCase();
      return feFirstName === originalFirstName || feFullNameLower === originalFeeEarnerFromEntry.toLowerCase();
    });
    
    if (matchingFeeEarner) {
      originalFeeEarnerFullName = matchingFeeEarner.name;
    }
    
    // Replace missing fee earner name with original fee earner's FULL name in narrative
    let swappedNarrative = originalNarrative;
    
    // First try to replace the full name if present
    swappedNarrative = swappedNarrative.replace(
      new RegExp(missingFeeEarnerName, "gi"),
      originalFeeEarnerFullName
    );
    
    // Then try to replace just the first name with the full name
    const missingFirstName = missingFeeEarnerName.split(" ")[0];
    // Use word boundaries to ensure we're replacing whole words
    swappedNarrative = swappedNarrative.replace(
      new RegExp(`\\b${missingFirstName}\\b`, "gi"),
      originalFeeEarnerFullName
    );
    
    return swappedNarrative;
  };
  
  // Test cases
  const tests = [
    {
      desc: "Replace first name only with full name",
      originalNarrative: "Call with Sophie Smith",
      missingFeeEarnerName: "Sophie Smith",
      originalFeeEarnerFromEntry: "Callum",
      expected: "Call with Callum Reyes"
    },
    {
      desc: "Replace full name with full name",
      originalNarrative: "Meeting with Sophie Smith about case",
      missingFeeEarnerName: "Sophie Smith",
      originalFeeEarnerFromEntry: "William Chen",
      expected: "Meeting with William Chen about case"
    },
    {
      desc: "Multiple occurrences of first name",
      originalNarrative: "Call with Sophie and Sophie's team",
      missingFeeEarnerName: "Sophie Smith",
      originalFeeEarnerFromEntry: "Theo",
      expected: "Call with Theo Johnson and Theo Johnson's team"
    },
    {
      desc: "Case insensitive replacement",
      originalNarrative: "CALL WITH SOPHIE SMITH",
      missingFeeEarnerName: "Sophie Smith",
      originalFeeEarnerFromEntry: "callum",
      expected: "CALL WITH Callum Reyes"
    },
    {
      desc: "Only replace whole words",
      originalNarrative: "Meeting with Sophie about philosophy",
      missingFeeEarnerName: "Sophie Smith",
      originalFeeEarnerFromEntry: "William",
      expected: "Meeting with William Chen about philosophy"
    },
    {
      desc: "Original fee earner not in list (use as-is)",
      originalNarrative: "Call with Sophie Smith",
      missingFeeEarnerName: "Sophie Smith",
      originalFeeEarnerFromEntry: "Unknown Person",
      expected: "Call with Unknown Person"
    }
  ];
  
  tests.forEach(test => {
    const result = swapNamesInNarrative(
      test.originalNarrative,
      test.missingFeeEarnerName,
      test.originalFeeEarnerFromEntry,
      feeEarners
    );
    const status = result === test.expected ? '✓' : '✗';
    console.log(`${status} ${test.desc}`);
    console.log(`   Input: "${test.originalNarrative}" (from ${test.originalFeeEarnerFromEntry})`);
    console.log(`   Output: "${result}"`);
    if (result !== test.expected) {
      console.log(`   Expected: "${test.expected}"`);
    }
    console.log('');
  });
}

// Run the test
console.log('=== Running Name Swapping Tests ===');
testNameSwapping();
console.log('\n=== Tests Complete ===');
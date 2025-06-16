// Test via clasp-gas-unit
function fetchPtsTest() {
  var data = fetchPts();
  console.log('Fetched data length:', data.length);
  if (data.length > 0) {
    console.log('First item:', JSON.stringify(data[0]));
  }
}
const styles = {
  headerDark: {
    fill: {
      fgColor: {
        rgb: 'FF008000'
      }
    },
    font: {
      color: {
        rgb: 'FFFFFFFF'
      },
      sz: 18,
      bold: true
    }
  },
  title: {
    fill: {
      fgColor: {
        rgb: 'FFE0E0E0'
      },
    },
    font: {
      color: {
        rgb: 'FF0080C4'
      },
      sz: 34
    }
  },
  data: {
    font: {
      sz: 16
    }
  }
};

//Array of objects representing heading rows (very top)
const heading = [
  [{ value: 'This is the Title', style: styles.title }] // <-- It can be only values
];

//Here you specify the export structure
const specification = {
  customer_name: { // <- the key should match the actual data key
    displayName: 'Customer', // <- Here you specify the column header
    headerStyle: styles.headerDark, // <- Header style
    //cellStyle: function (value, row) { // <- style renderer function
      // if the status is 1 then color in green else color in red
      // Notice how we use another cell value to style the current one
      //return (row.status_id == 1) ? styles.cellGreen : { fill: { fgColor: { rgb: 'FFFF0000' } } }; // <- Inline cell style is possible 
    //},
    cellStyle: styles.data,
    width: 120 // <- width in pixels
  },
  status_id: {
    displayName: 'Status',
    headerStyle: styles.headerDark,
    //cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
    //  return (value == 1) ? 'Active' : 'Inactive';
    //},
    cellStyle: styles.data,
    width: '10' // <- width in chars (when the number is passed as string)
  },
  note: {
    displayName: 'Description',
    headerStyle: styles.headerDark,
    cellStyle: styles.data, // <- Cell style
    width: 220 // <- width in pixels
  }
}

// The data set should have the following shape (Array of Objects)
// The order of the keys is irrelevant, it is also irrelevant if the
// dataset contains more fields as the report is build based on the
// specification provided above. But you should have all the fields
// that are listed in the report specification
const dataset = [
  { customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown' },
  { customer_name: 'HP', status_id: 0, note: 'some note' },
  { customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown' }
];

// Define an array of merges. 1-1 = A:1
// The merges are independent of the data.
// A merge will overwrite all data _not_ in the top-left cell.
const merges = [
  { start: { row: 1, column: 1 }, end: { row: 1, column: 3 } }
];


const schema = {
  'Customer': {
    prop: 'customer',
    type: String
    // Excel stores dates as integers.
    // E.g. '24/03/2018' === 43183.
    // Such dates are parsed to UTC+0 timezone with time 12:00 .
  },
  'Status': {
    prop: 'status',
    type: String,
    required: true
  },
  // 'COURSE' is not a real Excel file column name,
  // it can be any string — it's just for code readability.
  'Description': {
    prop: 'course',
    type: String
  }
}

module.exports = { styles, heading, specification, dataset, merges, schema };

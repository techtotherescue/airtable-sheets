/**
 * @fileoverview utils.js tests.
 */


const utils = require('./utils');


it.each`
    input               | output
    ${[{a: 1}]}         | ${["a"]}
    ${[{a: 1}, {b:2}]}  | ${["a", "b"]}
`('Extracting $input should return $output', ({input, output}) => {
  expect(utils.getHeaderFromRecords(input)).toEqual(output);
});


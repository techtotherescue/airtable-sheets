/**
 * @fileoverview sheet.js tests.
 */


const sheet = require('./sheet');


it.each`
    input               | output
    ${[1, 2, 3]}        | ${[[1, 3]]}
    ${[1, 2, 4, 5]}     | ${[[1, 2], [4, 2]]}
    ${[3, 4, -1, 0, 8]} | ${[[3, 2], [-1, 2], [8, 1]]}
`('Batching $input should return $output', ({input, output}) => {
  expect(sheet.batchConsecutiveIntegers(input)).toEqual(output);
});

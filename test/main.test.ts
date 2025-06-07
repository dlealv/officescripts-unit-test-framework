// test/main.test.ts
import { mainFn, TestCaseClass } from '../src/wrapper';

describe('Office Script tests', () => {
  it('main runs without error', () => {
    const mockWorkbook = {}; // Your mock implementation here
    mainFn(mockWorkbook);
  });

  it('TestCase can run test', () => {
    const testCase = new TestCaseClass();
    testCase.test('dummy test', () => {
      // your assertions
    });
    testCase.run();
  });
});

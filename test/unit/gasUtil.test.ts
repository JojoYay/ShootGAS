/**
 * Unit tests for GasUtil with mocked GasProps.
 * GasProps.instance.mappingSheet is mocked to return fixed data.
 */
jest.mock('../../src/gasProps', () => ({
  GasProps: {
    instance: {
      mappingSheet: {
        getDataRange: () => ({
          getValues: () => [
            ['lineName', 'densukeName', 'userId', 'role'],
            ['line1', 'name1', 'Ukanji1', '幹事'],
            ['line2', 'name2', 'UnotKanji', 'member'],
          ],
        }),
      },
    },
  },
}));

import { GasUtil } from '../../src/gasUtil';

describe('GasUtil', () => {
  describe('isKanji', () => {
    it('returns true when userId is in kanji role', () => {
      const gasUtil = new GasUtil();
      expect(gasUtil.isKanji('Ukanji1')).toBe(true);
    });

    it('returns false when userId is not kanji', () => {
      const gasUtil = new GasUtil();
      expect(gasUtil.isKanji('UnotKanji')).toBe(false);
    });

    it('returns false for unknown userId', () => {
      const gasUtil = new GasUtil();
      expect(gasUtil.isKanji('Uunknown')).toBe(false);
    });

    it('returns false for empty string', () => {
      const gasUtil = new GasUtil();
      expect(gasUtil.isKanji('')).toBe(false);
    });
  });
});

/**
 * Copyright 2024 JojoYay
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/**
 * Unit tests for GasUtil with mocked GasProps.
 * GasProps.instance.usersSheet is mocked to return fixed data.
 * For "usersSheet not set, fallback to DensukeMapping" behavior, see gasProps.test.ts.
 */
jest.mock('../../src/gasProps', () => ({
    GasProps: {
        instance: {
            usersSheet: {
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

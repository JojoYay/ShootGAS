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
 * Unit tests for GasProps.usersSheet fallback:
 * When usersSheet property is not set, the getter should use settingSheet and sheet "DensukeMapping".
 */
const DENSUKE_LEGACY_FIXTURE = [
    ['lineName', 'densukeName', 'userId', 'role'],
    ['line1', 'name1', 'Ulegacy1', '幹事'],
    ['line2', 'name2', 'Ulegacy2', 'member'],
];

jest.mock('../../src/scriptProps', () => ({
    ScriptProps: {
        instance: {
            usersSheet: 'setting-sheet-id',
            settingSheet: 'setting-sheet-id',
        },
    },
}));

beforeAll(() => {
    (global as unknown as { SpreadsheetApp: unknown }).SpreadsheetApp = {
        // openById: (id: string) => ({
        openById: () => ({
            getSheetByName: (name: string) => {
                if (name === 'DensukeMapping') {
                    return {
                        getDataRange: () => ({
                            getValues: () => DENSUKE_LEGACY_FIXTURE,
                        }),
                    };
                }
                return null;
            },
        }),
    };
});

import { GasProps } from '../../src/gasProps';

describe('GasProps.usersSheet', () => {
    describe('when usersSheet property is not set (same as settingSheet)', () => {
        it('returns sheet from settingSheet with name DensukeMapping', () => {
            const sheet = GasProps.instance.usersSheet;
            expect(sheet).toBeDefined();
        });

        it('returns data from DensukeMapping sheet', () => {
            const sheet = GasProps.instance.usersSheet;
            const values = sheet.getDataRange().getValues();
            expect(values).toEqual(DENSUKE_LEGACY_FIXTURE);
        });
    });
});

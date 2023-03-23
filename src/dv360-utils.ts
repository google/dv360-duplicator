/**
    Copyright 2023 Google LLC
    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at
        https://www.apache.org/licenses/LICENSE-2.0
    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.
*/

import { DV360 } from "./dv360";

export type LogicalOperation = 'OR'|'AND';

export const DV360Utils = {
    generateFilterString(
        field: string,
        values: string[],
        logicalOperation: LogicalOperation = 'OR'
    ) {
        if (! field) {
            throw Error('Filtering field canot be empty');
        }

        if (!values.length) {
            return '';
        }

        return '(' 
            + values
                .map((filterValue) => `${field}="${filterValue}"`)
                .join(` ${logicalOperation} `)
            + ')';

    },
};

export const dv360 = new DV360(ScriptApp.getOAuthToken());
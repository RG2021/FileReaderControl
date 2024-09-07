/* App.tsx */

import * as React from 'react';
import { IInputs } from "./generated/ManifestTypes";
import { DetailsList, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import { useState, useEffect } from 'react';
import {detailsListStyles} from './ControlStyles';

export function PCFControl(context : ComponentFramework.Context<IInputs>) {

    async function fetchFileData(entityName: string, recordId: string, fileColumn: string) 
    {
        // API endpoint to get data from file-type column
        // Example: https://org*******.crm8.dynamics.com/api/data/v9.2/accounts(cd961e9a-a83b-****-****-*****)/new_file/$value

        const apiURL = `${window.location.origin}/api/data/v9.1/${entityName}(${recordId})/${fileColumn}/$value`;

        const response = await fetch(apiURL, {
            method: 'GET',
            headers: {
                'OData-Version': '4.0',
                'OData-MaxVersion': '4.0',
                'Accept': 'application/octet-stream'
            },
            credentials: 'same-origin'
        });

        // File Exists
        if (response.status === 200) {
            const blob = await response.blob(); // File Binary Data to Blob
            const result = await new Response(blob).text(); // Blob to Text
            return JSON.parse(result); // Text to JSON
        } 
        
        return [];
    }

    const [records, setRecords] = useState<Array<Record<string, any>>>([]);

    useEffect(() => {

        // Retrive property values using context
        const entityName = context.parameters.entityName.raw || "";
        const recordID = context.parameters.recordID.raw || "";
        const fileColumn = context.parameters.fileColumn.raw || "";

        fetchFileData(entityName, recordID, fileColumn).then((data) => {
            setRecords(data); // Updates the records variable with new data
        });

    }, []);

    return (
        <DetailsList 
            items={records}
            styles={detailsListStyles}
            data-is-scrollable={false}
            layoutMode={DetailsListLayoutMode.fixedColumns} 
        />
    );
}
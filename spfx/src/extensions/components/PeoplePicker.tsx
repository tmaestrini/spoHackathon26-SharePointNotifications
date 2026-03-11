import { IItemPickerOption, IItemPickerProps, ItemPicker, useApplicationContext } from '@spteck/react-controls-v2';
import * as React from 'react';


export interface IPeoplePickerProps extends Omit<IItemPickerProps, 'options' | 'onSearchChange' | 'selectedOptions' | 'onSelectionChange'> {
    defaultSelectedIds?: string[];
    onPeopleChange?: (selectedIds: string[]) => void;
}

export const PeoplePicker: React.FunctionComponent<IPeoplePickerProps> = (props: React.PropsWithChildren<IPeoplePickerProps>) => {
    const [options, setOptions] = React.useState<IItemPickerOption[]>([]);
    const [selectedOptions, setSelectedOptions] = React.useState<IItemPickerOption[]>([]);
    const context = useApplicationContext();

    const FetchData: (searchParams: string) => Promise<IItemPickerOption[]> = async (searchParams: string) => {
        if (!context?.graphClient) return [];
        const response: { value: { id: string; displayName: string; mail: string }[] } = await context.graphClient.get(searchParams);
        const users: IItemPickerOption[] = response.value.map((user) => ({
            text: user.displayName,
            value: user.id
        }));
        return users ?? [];
    }


    const PerformSearch = async (searchText?: string) => {
        const url = new URL(`/users`, "http://localhost");
        url.searchParams.set("$filter", `accountEnabled eq true`)
        url.searchParams.set("$top", "10");
        url.searchParams.set("$select", "id,displayName,mail");
        if (searchText) {
            url.searchParams.set("$filter", `accountEnabled eq true and startswith(displayName,'${searchText}')`);

            // Using search requires ConsistencyLevel: eventual header, @spteck/react-controls-v2 doesn't support setting custom headers
            // url.searchParams.set("$search", `DisplayName:${searchText}*`);
        }
        const users = await FetchData(url.pathname + url.search);
        setOptions(users);
    }



    React.useEffect(() => {

        const loadInitialSelections = async () => {
            const url = new URL(`/users`, "http://localhost");
            url.searchParams.set("$select", "id,displayName,mail");
            url.searchParams.set("$filter", `accountEnabled eq true and (${props?.defaultSelectedIds?.map(option => `id eq '${option}'`).join(' or ')})`);
            const selection = await FetchData(url.pathname + url.search);
            if (selection)
                setSelectedOptions(selection);
        }

        if (props.defaultSelectedIds && props.defaultSelectedIds.length > 0) {
            loadInitialSelections()
                .catch(error => console.error('Error loading initial selections:', error));
        }

        PerformSearch()
            .catch(error => console.error('Error performing initial search:', error));

    }, []);



    return (
        <ItemPicker
            options={options}
            onSearchChange={(val) => {
                PerformSearch(val)
                    .catch(error => console.error('Error performing search:', error));
            }}
            selectedOptions={selectedOptions}
            onSelectionChange={(selectedIds) => {
                setSelectedOptions(selectedIds);
                if (props.onPeopleChange) {
                    props.onPeopleChange(selectedIds.map(option => option.value));
                }
            }}
            {...props}
        />
    );
};
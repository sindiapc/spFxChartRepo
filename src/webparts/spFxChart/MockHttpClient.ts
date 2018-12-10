import { IProjectItem } from './SpFxChartWebPart';

export default class MockHttpClient {
    private static _items: IProjectItem[] = [
        { Id: '1', Title: 'Project A', TeamSize: 2 },
        { Id: '2', Title: 'Project B', TeamSize: 20 },
        { Id: '3', Title: 'Project C', TeamSize: 10 },
    ];

    public static get(): Promise<IProjectItem[]> { 
        return new Promise<IProjectItem[]>((resolve)=>{
            resolve(MockHttpClient._items);
        })   ;    
        // return new Promise<IProjectItem[]>((resolve)=>{
        //     resolve(MockHttpClient._items);
        // });
    }
}
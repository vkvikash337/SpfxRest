import { ISPList } from './SharepointDataWebPart';  
  
export default class MockHttpClient {  
    private static _items: ISPList[] = [{ Title: 'E123'},];  
    public static get(restUrl: string, options?: any): Promise<ISPList[]> {  
      return new Promise<ISPList[]>((resolve) => {  
            resolve(MockHttpClient._items);  
        });  
    }  
}  
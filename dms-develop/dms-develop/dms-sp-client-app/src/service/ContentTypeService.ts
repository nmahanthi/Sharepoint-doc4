import { AadHttpClient } from "@microsoft/sp-http";

export interface ContentTypeResponse {
  displayName: string;
  id: string;
  template?: string;
}

export interface ContentTypeCreateRequest {
  displayName: string;
  template: string;
}
export interface ContentTypeUpdateRequest extends ContentTypeCreateRequest {
  parentId?: string;
}

export class ContentTypeService {
  private readonly baseUrl: string;
  private readonly siteUrl: string;
  private _dmsClient: AadHttpClient;
  constructor(baseUrl: string, siteUrl: string, dmsClient: AadHttpClient) {
    this.baseUrl = baseUrl;
    this.siteUrl = encodeURIComponent(siteUrl);
    this._dmsClient = dmsClient;

  }
  async getAllContentTypes(): Promise<ContentTypeResponse[]> {
    try {
      const response = await this._dmsClient.get(
        `${this.baseUrl}&siteUrl=${this.siteUrl}`,
        AadHttpClient.configurations.v1
      );
      if (!response.ok) {
        throw new Error(`Failed to fetch content types: ${response.statusText}`);
      }
      const data = await response.json();
      console.log("getAllContentTypes",data);
      
      return data;
    } catch (error) {
      console.error("Error fetching content types:", error);
      throw error;
    }
  }
  async addContentType(data: ContentTypeCreateRequest): Promise<ContentTypeResponse> {
    const httpTriggerEndPoint = `${this.baseUrl}&siteUrl=${this.siteUrl}`;
    try {
      const response = await this._dmsClient.post(httpTriggerEndPoint, AadHttpClient.configurations.v1, {
        body: JSON.stringify(data)
      });
      return response.json();
    } catch (error) {
      console.error("Error adding content type item", error);
      throw new Error("Error adding content type item");
    }
  }
  
  async deleteContentType(id: string): Promise<void> {
    const httpTriggerEndPoint = `${this.baseUrl}&siteUrl=${this.siteUrl}&id=${id}`;
    try {
      await this._dmsClient.fetch(httpTriggerEndPoint, AadHttpClient.configurations.v1, {
        method: 'DELETE',
        body: JSON.stringify({ })
      });
    } catch (error) {
      console.error("Error deleting content type item", error);
      throw new Error("Error deleting content type item");
    }
  }

  async updateContentType(id: string, data: ContentTypeCreateRequest): Promise<ContentTypeResponse> {
   const httpTriggerEndPoint = `${this.baseUrl}&siteUrl=${this.siteUrl}&id=${id}`;

    try {
        const response = await this._dmsClient.fetch(httpTriggerEndPoint, AadHttpClient.configurations.v1, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(data)
        });

        // Check if the response is successful
        if (!response.ok) {
            throw new Error(`Failed to update content type: ${response.statusText}`);
        }

        return await response.json(); // Parse the JSON response if the update is successful
    } catch (error) {
        console.error("Error updating content type item", error);
        throw new Error("Error updating content type item");
    }
}
  
}
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { MSGraphClientFactory } from "@microsoft/sp-http";

// GET
export const fetchData = async (graphClient: MSGraphClientFactory, query: string, reference: string): Promise<any> => {
  try {
    const client = await graphClient.getClient("3");
    const res = await client.api(query).get();
    return res;
  } catch (error) {
    console.log(reference, error);
    throw error;
  }
};

// POST
export const createResource = async (graphClient: MSGraphClientFactory, query: string, content: any, reference: string) => {
  try {
    const client = await graphClient.getClient("3");
    const response = await client.api(query)
      .header("Content-type", "application/json")
      .post(content);

    return response;
  } catch (error) {
    console.log(reference, error);
  }
};

// PATCH - Pass etag as IF-Match value
export const updateResource = async (graphClient: MSGraphClientFactory, query: string, etag: string, content: any, reference: string) => {
  try {
    const client = await graphClient.getClient("3");

    if (etag) {
      try {
        const response = await client.api(query)
          .headers({ "Content-type": "application/json", "If-Match": etag })
          .patch(content);
        return response;
      } catch (err) {
        console.log(reference, err);
        throw err; // Re-throw the error for handling outside the async function
      }
    } else {
      try {
        const response = await client.api(query)
          .header("Content-type", "application/json")
          .patch(content);
        return response;
      } catch (err) {
        console.log(reference, err);
        throw err; // Re-throw the error for handling outside the async function
      }
    }
  } catch (error) {
    console.log(reference, error);
  }
};

// DELETE - Pass IF-Match value as etag
export const deleteResource = async (graphClient: MSGraphClientFactory, query: string, etag: string, reference: string) => {
  try {
    const client = await graphClient.getClient("3");
    const response = await client.api(query)
      .header("If-Match", etag)
      .delete();
    return response;
  } catch (error) {
    console.log(reference, error);
  }
};

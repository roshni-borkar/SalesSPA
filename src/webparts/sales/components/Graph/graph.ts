/* eslint-disable promise/param-names */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable dot-notation */
/* eslint-disable @typescript-eslint/no-unused-expressions */
import * as _ from "lodash";
import { MSGraphClientFactory } from "@microsoft/sp-http";
import { createResource, deleteResource, fetchData, updateResource } from "./graphFunctions";
// import type { TCustomFieldType } from "../../Types/GlobalTypes";
// import { TPlannerDetailsObj, TPlannerObj } from "../Types/GanttTypes";

class Graph {
  private graphClient: MSGraphClientFactory;

  constructor(graphClient: MSGraphClientFactory) {
    this.graphClient = graphClient;
  }

  private blobToBase64(blob: Blob) {
    return new Promise((resolve, _) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.readAsDataURL(blob);
    });
  }

  private extractContent(html: string) {
    const doc = new DOMParser()
      .parseFromString(html, "text/html");

    // removing everything after the first table from the comment
    const table = doc.querySelector('table');
    if (table) {
      let nextSibling = table.nextSibling;
      while (nextSibling) {
        const toRemove = nextSibling;
        nextSibling = nextSibling.nextSibling;
        if (toRemove.parentNode) {
          toRemove.parentNode.removeChild(toRemove);
        }
      }
      if (table.parentNode) {
        table.parentNode.removeChild(table);
      }
    }

    // removing the p tag after the "Disclaimer"
    const pTag = doc.querySelector('p');
    pTag && pTag.remove();

    // returning the comment's text content
    return doc.documentElement?.textContent?.replace("Disclaimer", "") || "";
  }

  // public getColumnContent(obj: { type: TCustomFieldType, name: string; displayName: string; choiceValues: string; isRequired: boolean; }) {
  //   let columnContent = {
  //     "description": "",
  //     "enforceUniqueValues": false,
  //     "hidden": false,
  //     "indexed": false,
  //     "name": obj.name,
  //     "displayName": obj.displayName,
  //     "required": obj.isRequired
  //   };

  //   if (obj.type === "text") {
  //     columnContent["text"] = {
  //       "allowMultipleLines": false,
  //       "appendChangesToExistingText": false,
  //       "linesForEditing": 0,
  //       "maxLength": 255
  //     };
  //   } else if (obj.type === "number") {
  //     columnContent["number"] = {
  //       "decimalPlaces": "none",
  //       "displayAs": "number"
  //     };
  //   } else if (obj.type === "dateTime") {
  //     columnContent["dateTime"] = {
  //       "displayAs": "standard",
  //       "format": "dateOnly"
  //     };
  //   } else if (obj.type === "choice") {
  //     columnContent["choice"] = {
  //       "choices": obj.choiceValues.split(",").map((item) => item.trim()),
  //       "displayAs": "dropDownMenu"
  //     };
  //   }

  //   return columnContent;
  // };

  // public async sendMail(content) {
  //   return await createResource(this.graphClient, `/me/sendMail`, content, "send mail");
  // }

  public async getGroupRelatedSiteDetails(groupId: string) {
    return await fetchData(this.graphClient, `/groups/${groupId}/sites/root`, "get site details");
  }

  public async getSiteLists(groupId: string) {
    const siteDetails = await this.getGroupRelatedSiteDetails(groupId);
    return (await fetchData(this.graphClient, `/sites/${siteDetails.id}/lists`, "lists")).value;
  }

  public async createList(groupId: string, { displayName, columns }: { displayName: string; columns?: any[]; }) {
    const siteDetails = await this.getGroupRelatedSiteDetails(groupId);
    return await createResource(this.graphClient, `/sites/${siteDetails.id}/lists`, displayName && columns ? { displayName, columns } : { displayName }, `Created ${displayName}`);
  }

  public async getListColumns(groupId: string, listId: string) {
    const siteDetails = await this.getGroupRelatedSiteDetails(groupId);
    return (await fetchData(this.graphClient, `/sites/${siteDetails.id}/lists/${listId}/columns`, 'get list columns')).value;
  }

  public async createColumnDefinition(groupId: string, listId: string, columnContent: any) {
    const siteDetails = await this.getGroupRelatedSiteDetails(groupId);
    return await createResource(this.graphClient, `/sites/${siteDetails.id}/lists/${listId}/columns`, columnContent, "Create a column");
  }

  public async updateColumnDefinition(groupId: string, listId: string, columnName: string, updateData: any) {
    const siteDetails = await this.getGroupRelatedSiteDetails(groupId);
    const listColumns = await this.getListColumns(groupId, listId);
    const selectedColumn = listColumns?.find((col: { name: string }) => col.name === columnName);

    return await updateResource(this.graphClient, `/sites/${siteDetails.id}/lists/${listId}/columns/${selectedColumn.id}`, selectedColumn['@odata.etag'] || '', updateData, "updaing column definition");
  }
  public async deleteColumnDefinition(groupId: string, listId: string, columnName: string) {
    const siteDetails = await this.getGroupRelatedSiteDetails(groupId);
    const listColumns = await this.getListColumns(groupId, listId);
    const selectedColumn = listColumns?.find((col: { name: string }) => col.name === columnName);

    return await deleteResource(this.graphClient, `/sites/${siteDetails.id}/lists/${listId}/columns/${selectedColumn.id}`, selectedColumn['@odata.etag'] || '', "deleting column definition");
  }

  public async getListItems(groupId: string, listId: string) {
    const siteDetails = await this.getGroupRelatedSiteDetails(groupId);
    return (await fetchData(this.graphClient, `/sites/${siteDetails.id}/lists/${listId}/items`, "get list items")).value;
  }


  public async getGroupMembers(groupId: string) {
    return (await fetchData(this.graphClient, `/groups/${groupId}/members?$select=id,mail,displayName,accountEnabled,userPrincipalName`, "get all members of group")).value;
  }

  public async getGroupOwners(groupId: string) {
    return (await fetchData(this.graphClient, `/groups/${groupId}/owners`, "get all owners of group")).value;
  }

  public async getUserPhoto(mailId: string) {
    const blob = await fetchData(this.graphClient, `/users/${mailId?.toLowerCase()}/photo/$value`, "get user photo");
    return await this.blobToBase64(blob);
  }


  public async getConversationReplies(groupId: string, conversationThreadId: string) {
    const conversation = (await fetchData(this.graphClient, `/groups/${groupId}/conversations/${conversationThreadId}/threads/${conversationThreadId}/posts/`, "get task comments")).value;

    const conversationPosts = conversation.map((comment: { id: string; createdDateTime: string; sender: { emailAddress: string }; body: { content: string } }) => {
      const extractedcontent = this.extractContent(comment.body.content);
      const content = extractedcontent.split("\n")
        .filter(item => item.trim());

      return {
        id: comment.id,
        createdDateTime: comment.createdDateTime,
        sender: comment.sender.emailAddress,
        content
      };
    }).reverse();

    return conversationPosts;
  }

  public async createConversationReply(groupId: string, conversationThreadId: string, content: string) {
    const reply = {
      post: {
        body: {
          contentType: 'html',
          content
        },
        newParticipants: []
      },
    };

    return await createResource(this.graphClient, `/groups/${groupId}/conversations/${conversationThreadId}/threads/${conversationThreadId}/reply`, reply, "create a reply");
  }
}

export default Graph;
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpPnPjscrudWebPart.module.scss';
import * as strings from 'SpPnPjscrudWebPartStrings';

import { IListItem } from './IListItem';
import pnp, { sp, Item, ItemAddResult, ItemUpdateResult } from "sp-pnp-js";

export interface ISpPnPjscrudWebPartProps {
  listName: string;
}

export default class SpPnPjscrudWebPart extends BaseClientSideWebPart<ISpPnPjscrudWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spPnPjscrud }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">CRUD operations</span>
              <p class="${ styles.subTitle }">SP PnP JS</p>
              <p class="${ styles.description }">Name: ${escape(this.properties.listName)}</p>

              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <button class="${styles.button} create-Button">
                    <span class="${styles.label}">Create item</span>
                  </button>
                  <button class="${styles.button} read-Button">
                    <span class="${styles.label}">Read item</span>
                  </button>
                </div>
              </div>

              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <button class="${styles.button} update-Button">
                    <span class="${styles.label}">Update item</span>
                  </button>
                  <button class="${styles.button} delete-Button">
                    <span class="${styles.label}">Delete item</span>
                  </button>
                </div>
              </div>

              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                  <div class="status"></div>
                  <ul class="items"><ul>
                </div>
              </div>              

            </div>
          </div>
        </div>
      </div>`;

      this.setButtonsEventHandlers();
  }

  private setButtonsEventHandlers(): void {
    const webPart: SpPnPjscrudWebPart = this;
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.createItem(); });
    this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.readItem(); });
    this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); });
    this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.deleteItem(); });
  }

  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      sp.web.lists.getByTitle(this.properties.listName)
        .items.orderBy('Id', false).top(1).select('Id').get()
        .then((items: { Id: number }[]): void => {
          if (items.length === 0) {
            resolve(-1);
          }
          else {
            resolve(items[0].Id);
          }
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  private updateStatus(status: string, items: IListItem[] = []): void {
    this.domElement.querySelector('.status').innerHTML = status;
    this.updateItemsHtml(items);
  }

  private updateItemsHtml(items: IListItem[]): void {
    this.domElement.querySelector('.items').innerHTML = items.map(item => `<li>${item.Title} (${item.Id})</li>`).join("");
  }

  private createItem(): void {
    this.updateStatus('Creating item...');

    sp.web.lists.getByTitle(this.properties.listName).items.add({
      'Title': `Item ${new Date()}`
    }).then((result: ItemAddResult): void => {
      const item: IListItem = result.data as IListItem;
      this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
    }, (error: any): void => {
      this.updateStatus('Error while creating the item: ' + error);
    });
  }

  private readItem(): void {
    this.updateStatus('Reading latest items...');

    this.getLatestItemId()
      .then((itemId: number): Promise<IListItem> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        this.updateStatus(`Loading information about item ID: ${itemId}...`);
        return sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(itemId).select('Title', 'Id').get();
      })
      .then((item: IListItem): void => {
        this.updateStatus(`Item ID: ${item.Id}, Title: ${item.Title}`);
      }, (error: any): void => {
        this.updateStatus('Loading latest item failed with error: ' + error);
      });
  }

  private updateItem(): void {
    this.updateStatus('Loading latest items...');
    let latestItemId: number = undefined;
    let etag: string = undefined;

    this.getLatestItemId()
      .then((itemId: number): Promise<Item> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.updateStatus(`Loading information about item ID: ${itemId}...`);
        return sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(itemId).get(undefined, {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata'
            }
          });
      })
      .then((item: Item): Promise<IListItem> => {
        etag = item["odata.etag"];
        return Promise.resolve((item as any) as IListItem);
      })
      .then((item: IListItem): Promise<ItemUpdateResult> => {
        return sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(item.Id).update({
            'Title': `Updated Item ${new Date()}`
          }, etag);
      })
      .then((result: ItemUpdateResult): void => {
        this.updateStatus(`Item with ID: ${latestItemId} successfully updated`);
      }, (error: any): void => {
        this.updateStatus('Loading latest item failed with error: ' + error);
      });
  }

  private deleteItem(): void {
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }

    this.updateStatus('Loading latest items...');
    let latestItemId: number = undefined;
    let etag: string = undefined;
    this.getLatestItemId()
      .then((itemId: number): Promise<Item> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        latestItemId = itemId;
        this.updateStatus(`Loading information about item ID: ${latestItemId}...`);
        return sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(latestItemId).select('Id').get(undefined, {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata'
            }
          });
      })
      .then((item: Item): Promise<IListItem> => {
        etag = item["odata.etag"];
        return Promise.resolve((item as any) as IListItem);
      })
      .then((item: IListItem): Promise<void> => {
        this.updateStatus(`Deleting item with ID: ${latestItemId}...`);
        return sp.web.lists.getByTitle(this.properties.listName)
          .items.getById(item.Id).delete(etag);
      })
      .then((): void => {
        this.updateStatus(`Item with ID: ${latestItemId} successfully deleted`);
      }, (error: any): void => {
        this.updateStatus(`Error deleting item: ${error}`);
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

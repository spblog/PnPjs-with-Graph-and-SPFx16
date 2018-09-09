import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphPnPWebpartWebPart.module.scss';
import * as strings from 'GraphPnPWebpartWebPartStrings';
import { GraphTokenFetchClient } from "../../GraphTokenFetchClient";
import { graph } from "@pnp/graph";
import { AadTokenProvider } from '@microsoft/sp-http';

export interface IGraphPnPWebpartWebPartProps {
  description: string;
}

export default class GraphPnPWebpartWebPart extends BaseClientSideWebPart<IGraphPnPWebpartWebPartProps> {

  public onInit(): Promise<void> {
    return new Promise((resolve, reject) => {
      this.context.aadTokenProviderFactory
        .getTokenProvider()
        .then((tokenProvider: AadTokenProvider) => {

          graph.setup({
            graph: {
              fetchClientFactory: () => {
                return new GraphTokenFetchClient(tokenProvider);
              }
            }
          });

          resolve();
        })
        .catch(reject);
    });
  }

  public render(): void {

    graph.groups.get()
      .then((data) => {
        console.log(data);
      });

    this.domElement.innerHTML = `
      <div class="${ styles.graphPnPWebpart}">
        hello world
      </div>`;
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

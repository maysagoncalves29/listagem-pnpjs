import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './styles.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  users: any[];

  public render(): void {
    console.log("get users");
    this.getUsers()
      .then((users: any[]) => {
        this.users = users;

        const userListHtml = this.renderUserList(users);

        this.domElement.innerHTML = `
          <div class="${styles}"></div>
          <div class="formulario">
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" />
            <form></form>
            ${userListHtml}
          </div>
        `;

        const updateButtons = document.querySelectorAll('.update-button');
        updateButtons.forEach((button) => {
          button.addEventListener('click', () => {
            const userId = parseInt(button.getAttribute('data-userid'));
            this.redirectToUpdatePage(userId);
          });
        });

        const deleteButtons = document.querySelectorAll('.delete-button');
        deleteButtons.forEach((button, index) => {
          button.addEventListener('click', () => {
            const userId = users[index].ID;
            this.deleteUser(userId);
          });
        });

        const addButton = document.querySelector('.add-button');
        addButton.addEventListener('click', () => {
          const formAluno = `${this.context.pageContext.web.absoluteUrl}/SitePages/CRUD-PnP---Finalizado.aspx`;
          window.location.href = formAluno;
        });
      })
      .catch((error: any) => {
        console.log(error);
      });
  }

  private redirectToUpdatePage(userId: number): void {
    const formUrl = `${this.context.pageContext.web.absoluteUrl}/SitePages/CRUD-PnP---Finalizado.aspx?itemId=${encodeURIComponent(userId)}`;
    window.location.href = formUrl;
  }

  private getUsers(): Promise<any[]> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Students')/items?$select=ID,Name`;

    return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          throw new Error(`Unable to fetch users. Error: ${response.statusText}`);
        }
      })
      .then((data: any) => {
        return data.value;
      });
  }

  private deleteUser(userId: number): void {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Students')/items(${userId})`;

    this.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    }).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        alert('Usuário Excluído');
        console.log('Usuário excluído com sucesso');
      } else {
        alert('erro');
        console.log('Erro ao excluir usuário');
      }
    }).catch((error: any) => {
      console.log(error);
    });
  }

  private renderUserList(users: any[]): string {
    let userListHtml = '';

    userListHtml += `
      <div class="user-list">
        <table class="table-format">
          <tr>
            <th>Name</th>
            <th>ID</th>
            <th>Ações</th>
          </tr>
    `;

    for (const user of users) {
      userListHtml += `
        <tr>
          <td>${user.Name}</td>
          <td>${user.ID}</td>
          <td>
            <button type="button" class="custom-button update-button" data-userid="${user.ID}">
              <i class="fas fa-pencil-alt"></i>
            </button>
            <button type="button" class="custom-button delete-button">
              <i class="fas fa-trash"></i>
            </button>
          </td>
        </tr>
      `;
    }

    userListHtml += `
        <tr>
          <td colspan="3">
            <button type="button" class="custom-button add-button">
              <i class="fas fa-plus"></i> Cadastrar Aluno
            </button>
          </td>
        </tr>
      </table>
    </div>
    `;

    return userListHtml;
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

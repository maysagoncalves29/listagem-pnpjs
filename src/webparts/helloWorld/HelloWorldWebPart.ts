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
  this.getUsers().then((users: any[]) => {
    // amazena a lista de usuários
    this.users = users;
    const userListContainer = document.createElement('div');
    userListContainer.classList.add('user-list');

    for (const user of users) {
      const userContainer = document.createElement('div');
      userContainer.classList.add('user-item');

      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.value = user.Name;
      userContainer.appendChild(nameInput);

      const idInput = document.createElement('input');
      idInput.type = 'text';
      idInput.value = user.ID;
      userContainer.appendChild(idInput);

      const buttonsContainer = document.createElement('div');
      buttonsContainer.classList.add('buttons-container');

      const updateButton = document.createElement('button');
      updateButton.type = 'button';
      updateButton.innerHTML = '<i class="fas fa-pencil-alt"></i> Atualizar Aluno';
      updateButton.addEventListener('click', () => {
        this.updateUser(user.ID, nameInput.value);
      });
      buttonsContainer.appendChild(updateButton);

      const deleteButton = document.createElement('button');
      deleteButton.type = 'button';
      deleteButton.innerHTML = '<i class="fas fa-trash"></i> Remover Aluno';
      deleteButton.addEventListener('click', () => {
        this.deleteUser(user.ID);
      });
      buttonsContainer.appendChild(deleteButton);

      userContainer.appendChild(buttonsContainer);
      userListContainer.appendChild(userContainer);
    }

    this.domElement.appendChild(userListContainer);
  }).catch((error: any) => {
    console.log(error);
  });
    this.domElement.innerHTML = `
    <div class="${styles}"></div>
      <div class="formulario">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" />
        <form>
        </form>
      </div>
`;

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
  private updateUser(userId: number, newName: string): void {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Students')/items(${userId})`;

    const body = {
      Name: newName
  }
  this.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
    headers: {
      'Content-Type': 'application/json',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
    },
    body: JSON.stringify(body)
  }).then((response: SPHttpClientResponse) => {
    if (response.ok) {
      alert('Usuário atualizado!')
      console.log('Usuário atualizado com sucesso');
    } else {
      console.log('Erro ao atualizar usuário');
    }
  }).catch((error: any) => {
    console.log(error);
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
      alert('Usuário Excluído')
      console.log('Usuário excluído com sucesso');
    } else {
      console.log('Erro ao excluir usuário');
    }
  }).catch((error: any) => {
    console.log(error);
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

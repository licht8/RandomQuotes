import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './TestWebPart.module.scss';

export interface ITestWebPartProps {
  userInput: string;
}

export default class TestWebPart extends BaseClientSideWebPart<ITestWebPartProps> {

  private localStorageKey = 'quoteOfTheDay';
  private localStorageDayKey = 'quoteOfTheDayDay';

  public render(): void {
    const savedDay = localStorage.getItem(this.localStorageDayKey);
    const currentDay = new Date().getDate().toString();

    let content: string;

    if (savedDay !== currentDay) {
      const userInputLines = this.properties.userInput.split('\n');
      const randomIndex = Math.floor(Math.random() * userInputLines.length);
      const displayText = userInputLines[randomIndex] || '';

      content = `
        <div class="${styles.test} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
          <p><strong>Quote of the day</strong></p>
          <p>${displayText}</p>
        </div>`;

      localStorage.setItem(this.localStorageKey, displayText);
      localStorage.setItem(this.localStorageDayKey, currentDay);
    } else {
      const savedQuote = localStorage.getItem(this.localStorageKey) || 'No quote available';
      content = `
        <div class="${styles.test} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
          <p><strong>Quote of the day:</strong></p>
          <p>${savedQuote}</p>
        </div>`;
    }

    this.domElement.innerHTML = content;
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('userInput', {
                  label: 'Quotes to Shuffle:',
                  multiline: true,
                  rows: 4
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

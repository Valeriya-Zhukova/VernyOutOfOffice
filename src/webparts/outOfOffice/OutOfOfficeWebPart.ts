import { Version } from '@microsoft/sp-core-library';
import {
	IPropertyPaneChoiceGroupOption,
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { getIconClassName } from '@uifabric/styling';

import styles from './OutOfOfficeWebPart.module.scss';
import * as strings from 'OutOfOfficeWebPartStrings';

export interface IOutOfOfficeWebPartProps {
	description: string;
}

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISPLists {
	value: ISPList[];
}
export interface ISPList {
	ID: string;
	Title: string;
	From: Date;
	To: Date;
	Type: string;
}

export default class OutOfOfficeWebPart extends BaseClientSideWebPart<IOutOfOfficeWebPartProps> {
	private _getListData(): Promise<ISPLists> {
		// const today: Date = new Date();
		const yesterday: Date = new Date();
		yesterday.setDate(yesterday.getDate() - 1);
		console.log(yesterday.toISOString());
		return this.context.spHttpClient
			.get(
				`https://vernycapital.sharepoint.com/_api/web/lists/GetByTitle('Нет в офисе')/Items?$select=ID,Title,From,To,Type&$orderby= Title asc&$filter=(From le datetime'${yesterday.toISOString()}') and (To ge datetime'${yesterday.toISOString()}')&$top=500`,
				SPHttpClient.configurations.v1
			)
			.then((response: SPHttpClientResponse) => {
				debugger;
				return response.json();
			});
	}

	// https://vernycapital.sharepoint.com/_api/web/lists/GetByTitle('%D0%9D%D0%B5%D1%82%20%D0%B2%20%D0%BE%D1%84%D0%B8%D1%81%D0%B5')/Items?$select=ID,Title,From,To,Type&$orderby=%20Title%20asc&$filter=(From%20le%20datetime%272022-04-04T18:00:00Z%27)%20and%20(To%20ge%20datetime%272022-04-04T18:00:00Z%27)

	private _renderListAsync(): void {
		this._getListData().then((response) => {
			this._renderList(response.value);
		});
	}

	private _renderList(items: ISPList[]): void {
		let html: string = `<ul>`;

		items.forEach((item: ISPList) => {
			const link = `https://vernycapital.sharepoint.com/Lists/OutOf/DispForm.aspx?ID=${item.ID}`;
			html += `
				<li>
					<div class="${styles.icon}">
						<i class="${getIconClassName(item.Type === 'Отпуск' ? 'World' : 'Arrivals')} ${
				item.Type === 'Отпуск' ? styles.vac : styles.arr
			}"></i>
						</div>
					<a target="_blank" href=${link}>${item.Title}</a>
				</li>
			`;
		});

		html += `</ul>`;

		const listContainer: Element = this.domElement.querySelector('#spListContainer');
		listContainer.innerHTML = html;

		console.log(items);
	}

	public render(): void {
		this.domElement.innerHTML = `
			<div class="${styles.outOfOffice}">
				<div class="${styles.container}">
					<div class="${styles.wrapper}">
						<div class="${styles.title}">
							<h1>
								<a target="_blank" href="https://vernycapital.sharepoint.com/Lists/OutOf/AllItems.aspx">
									Нет в офисе
								</a>
							</h1>
						</div>
						<div id="spListContainer" class="${styles.list}"></div>
					</div>
				</div>
			</div>`;
		this._renderListAsync();
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription,
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField('description', {
									label: strings.DescriptionFieldLabel,
								}),
							],
						},
					],
				},
			],
		};
	}
}

import MsgReader from '@kenjiuno/msgreader';
import {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription, NodeOperationError,
} from 'n8n-workflow';
import { decode } from 'base64-arraybuffer';

export class MsgParser implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'MsgParser',
		name: 'msgParser',
		icon: 'file:envelope-open-line-icon.svg',
		group: ['transform'],
		version: 1,
		description: 'Convert binary msg file to json object',
		defaults: {
			name: 'MsgParser',
		},
		inputs: ['main'],
		outputs: ['main'],
		properties: []
	};

	// The function below is responsible for actually doing whatever this node
	// is supposed to do. In this case, we're just appending the `myString` property
	// with whatever the user has entered.
	// You can make async calls and use `await`.
	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();

		items.forEach((item, itemIndex) => {
			try {
				if(item.binary) {
					const fileData = new MsgReader(decode(item.binary?.data?.data)).getFileData();
					item.json = {
						...fileData
					};
					item.binary = undefined;
				}else{
					item.json = {};
				}
			}catch (error) {
				if (this.continueOnFail()) {
					items.push({ json: this.getInputData(itemIndex)[0].json, error, pairedItem: itemIndex });
				} else {
					// Adding `itemIndex` allows other workflows to handle this error
					if (error.context) {
						// If the error thrown already contains the context property,
						// only append the itemIndex
						error.context.itemIndex = itemIndex;
						throw error;
					}
					throw new NodeOperationError(this.getNode(), error, {
						itemIndex,
					});
				}
			}
		});

		return [items];
	}
}

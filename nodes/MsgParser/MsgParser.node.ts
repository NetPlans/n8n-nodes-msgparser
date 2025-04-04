import MsgReader from '@kenjiuno/msgreader';
import {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription, NodeOperationError,
} from 'n8n-workflow';

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
		properties: [
			// Node properties which the user gets displayed and
			// can change on the node.
			{
				displayName: 'Input Binary Field',
				name: 'inputBinaryFieldName',
				type: 'string',
				default: 'data',
				placeholder: 'data',
				description: 'The name of the input binary field containing the file to be extracted',
			},
		],
	};

	// The function below is responsible for actually executing the node
	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				const item: INodeExecutionData = items[itemIndex];

				if(item.binary) {
					const fileData = new MsgReader(await this.helpers.getBinaryDataBuffer(0, (this.getNodeParameter('inputBinaryFieldName', itemIndex, 'data') as string))).getFileData();
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
		}

		return [items];
	}
}

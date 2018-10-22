## SPFx reusable React control field-picker-list-data

I develop a reusable control using @PnP/PnpJs, Office-ui-fabric-react control, that consist of select value as you type from a list column data. This is util when you have large list of items.

This reusable control allow search and select list column data and can be use in any SPFx webpart.


## Install Package

Install with :

npm install field-picker-list-data --save --save-exact

## Control Properties

<table style="width: 100%; height: 786px;">
<tbody>
<tr>
<th style="width: 220px;">Property</th>
<th>Type</th>
<th style="width: 80px;">Required</th>
<th>Description</th>
</tr>
<tr>
<td>listId</td>
<td>string</td>
<td>yes</td>
<td>Gui of List</td>
</tr>
<tr>
<td>columnInternalName</td>
<td>string</td>
<td>yes</td>
<td>InternalName of column to search and get values</td>
</tr>
<tr>
<td>onSelectedItem: (item:any) =>void;</td>
<td>function</td>
<td>yes</td>
<td>Callback function</td>
</tr>
<tr>
<td>className</td>
<td>string</td>
<td>no</td>
<td>CSS className</td>
</tr>
<tr>
<td>webUrl</td>
<td>string</td>
<td>no</td>
<td>URL of site if different of current site, user must have permissions</td>
</tr>
<tr>
<td>value</td>
<td>Array</td>
<td>no</td>
<td>Default Value</td>
</tr>
<tr>
<td>disabled</td>
<td>Boolean</td>
<td>no</td>
<td>Disable Control</td>
</tr>
<tr>
<td>itemLimit</td>
<td>number</td>
<td>yes</td>
<td>Number os items to select / return</td>
</tr>
<tr>
<td>context</td>
<td>WebPartContext|ApplicationCustomizerContext</td>
<td>yes</td>
<td>WebPart or Application customiser context</td>
</tr>
<tr>
<td>sugestedHeaderText</td>
<td>string</td>
<td>no</td>
<td>Text header to display</td>
</tr>
<tr>
<td>noresultsFoundTextstring</td>
<td>string</td>
<td>no</td>
<td>Text message when no items</td>
</tr>
</tbody>
</table>

<h2>Usage</h2>
Import  control:

<img class="alignnone size-full wp-image-50" src="https://jjm935611985.files.wordpress.com/2018/10/screenshot-2018-10-22-at-10-34-29.png" alt="Screenshot 2018-10-22 at 10.34.29" width="712" height="51" />

Set Properties :

<img class="alignnone size-full wp-image-51" src="https://jjm935611985.files.wordpress.com/2018/10/screenshot-2018-10-22-at-10-36-03.png" alt="Screenshot 2018-10-22 at 10.36.03" width="749" height="212" />

Sample :

<img class="alignnone size-full wp-image-52" src="https://jjm935611985.files.wordpress.com/2018/10/screenshot-2018-10-22-at-10-57-44.png" alt="Screenshot 2018-10-22 at 10.57.44" width="811" height="377" /><img class="alignnone size-full wp-image-53" src="https://jjm935611985.files.wordpress.com/2018/10/screenshot-2018-10-22-at-10-58-41.png" alt="Screenshot 2018-10-22 at 10.58.41" width="828" height="365" /><img class="alignnone size-full wp-image-54" src="https://jjm935611985.files.wordpress.com/2018/10/screenshot-2018-10-22-at-10-58-58.png" alt="Screenshot 2018-10-22 at 10.58.58" width="808" height="270" />

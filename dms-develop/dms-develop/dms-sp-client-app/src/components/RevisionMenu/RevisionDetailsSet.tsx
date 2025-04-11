
import * as React from 'react';
import {  DetailsList, DetailsListLayoutMode, IColumn,SelectionMode ,Link, Stack, } from '@fluentui/react';
import {Text} from '@fluentui/react'
export interface IRevisionDetailsSetProps {
  items: { title: string, modified: string, createdBy: string, projectRevision: string, link: string }[];
  onCancel: () => void;
}

const RevisionDetailsSet: React.FC<IRevisionDetailsSetProps> = ({ items }) => {
  console.log('Rendering SimpleDetailsTable component', { items });

  const columns: IColumn[] = [
    { key: 'column1', name: 'Document Link', fieldName: 'link', minWidth: 100, maxWidth: 200, isResizable: true, onRender: (item: any) => <Link href={item.link} target="_blank">{item.title}</Link> },
    { key: 'column2', name: 'Modified', fieldName: 'modified', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column3', name: 'Created By', fieldName: 'createdBy', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'column4', name: 'Project Revision', fieldName: 'projectRevision', minWidth: 100, maxWidth: 200, isResizable: true },
   
  ];

  return (
   <>

 <Stack 
 horizontalAlign='center'
 verticalAlign='center'
 styles={{
  root:{
    backgroundColor: '#163350d6',
    color: "white",
    marginTop: 14,
    height: 50
  } 

   
 }}
 >
<Text variant='large' as="h3" styles={{root:{color:'white',fontSize:22}}}>Revision(s)</Text>
 </Stack>
    <DetailsList
      items={items}
      columns={columns}
      setKey="set"
      layoutMode={DetailsListLayoutMode.justified}
      selectionMode={SelectionMode.none} 
      selectionPreservedOnEmptyClick={true}
    />
   </>
  );
};

export default RevisionDetailsSet;
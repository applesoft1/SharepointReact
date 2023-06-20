import * as React from 'react';
// import styles from './Fap.module.scss';
import { IFapProps } from './IFapProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { FC, useEffect, useState } from 'react';
import styles from './Fap.module.scss';
// import {
//   SPHttpClient,
//   SPHttpClientResponse
// } from '@microsoft/sp-http';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  id: string;
  title: string;
}

const list: ISPList[] = [
  {id: '1', title: 'Admin and Tools'},
  {id: '2', title: 'General'},
  {id: '3', title: 'Microsoft Viva Topics and Sharepoint'},
  {id: '4', title: 'Topic generation, curation, and discovery'},
]

const Fap: FC<IFapProps> = ({}) => {
  const [items, setItems] = useState<ISPList[]>([])
  useEffect(() => setItems(list), [])
  return (<div>
    <h1>My Test</h1>
    {items.map(item => <ul className={styles.list}>
      <li className={styles.listItem}>
        <span className="ms-font-l">{item.title}</span>
      </li>
    </ul>
    )}
  </div>)
}
export default Fap

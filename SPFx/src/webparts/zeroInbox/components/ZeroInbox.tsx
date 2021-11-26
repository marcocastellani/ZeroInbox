import * as React from "react";
import styles from "./ZeroInbox.module.scss";
import { IZeroInboxProps } from "./IZeroInboxProps";
import { Get, Person } from "@microsoft/mgt-react";
import { Stack, IStackTokens } from "office-ui-fabric-react";
import { ThemeProvider } from "@uifabric/foundation";
const innerStackTokens: IStackTokens = {
  childrenGap: 5,
  padding: 10,
};
const Riga = (props) => {
  const email = props.dataContext;

  return (
    <>
      <ThemeProvider>
        <Stack horizontal tokens={innerStackTokens}>
          <Stack.Item>
            <Person personQuery={email.sender.emailAddress.address}></Person>
          </Stack.Item>

          <Stack.Item>
            <a href={email.webLink} target="_blank" rel="noorigin">
              {email.subject}
            </a>
          </Stack.Item>
        </Stack>
      </ThemeProvider>
    </>
  );
};

export default class ZeroInbox extends React.Component<IZeroInboxProps, {}> {
  public render(): React.ReactElement<IZeroInboxProps> {
    return (
      <>
        <div className={styles.minibox}>
          <h1 className={styles.label1}>TODO</h1>
          <Get resource="/me/messages?$filter=categories/any(c:c eq 'TODO')&$select=sender,subject,webLink">
            <Riga template="value"></Riga>
          </Get>
        </div>
        <div className={styles.minibox}>
          <h1 className={styles.label1}>TOREAD</h1>
          <Get resource="/me/messages?$filter=categories/any(c:c eq 'TOREAD')&$select=sender,subject,webLink">
            <Riga template="value"></Riga>
          </Get>
        </div>
      </>
    );
  }
}

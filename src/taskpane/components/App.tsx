import * as React from "react";
import Header from "./Header";
import ChartInsertion from "./ChartInsertion";
import ExcelImport from "./ExcelImport";
import { makeStyles } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} />
      <ChartInsertion />
      <ExcelImport />
    </div>
  );
};

export default App;

import * as React from "react";
import { useRef, useState } from "react";
import { Button, Field, tokens, makeStyles } from "@fluentui/react-components";
import { insertRandomBarChart } from "../chartToSlide";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    marginTop: "20px",
  },
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: "10px",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    marginTop: "10px",
  },
});

const ChartInsertion: React.FC = () => {
  const styles = useStyles();
  const clickCountRef = useRef(0);
  const [error, setError] = useState<string | null>(null);

  const handleAddChartClick = async () => {
    try {
      setError(null);
      await insertRandomBarChart(clickCountRef.current++);
    } catch (e) {
      setError(`Failed to insert chart: ${e instanceof Error ? e.message : String(e)}`);
    }
  };

  return (
    <div className={styles.container}>
      <Field className={styles.instructions}>Click the button to add a D3 bar chart to the current slide.</Field>
      <Button appearance="primary" size="large" onClick={handleAddChartClick}>
        Add D3 Bar Chart
      </Button>
      {error && <div className={styles.error}>{error}</div>}
    </div>
  );
};

export default ChartInsertion;

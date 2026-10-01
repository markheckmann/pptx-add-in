import * as React from "react";
import { Image, tokens, makeStyles } from "@fluentui/react-components";

export interface HeaderProps {
  title: string;
  logo: string;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "8px",
    paddingTop: "24px",
    paddingBottom: "20px",
    backgroundColor: tokens.colorNeutralBackground3,
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const { title, logo } = props;
  const styles = useStyles();

  return (
    <section className={styles.header}>
      <Image width="48" height="48" src={logo} alt={title} />
      <h1 className={styles.title}>{title}</h1>
    </section>
  );
};

export default Header;

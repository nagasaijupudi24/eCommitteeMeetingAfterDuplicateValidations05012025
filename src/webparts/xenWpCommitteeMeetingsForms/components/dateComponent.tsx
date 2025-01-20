/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./XenWpCommitteeMeetingsForms.module.scss";

const DateTime: React.FC = () => {
  const [currentDate, setCurrentDate] = useState(new Date());

  useEffect(() => {
    const timerID = setInterval(() => setCurrentDate(new Date()), 1000);
    return () => clearInterval(timerID);
  }, []);

  const formatWithZero = (value: number) => value < 10 ? `0${value}` : `${value}`;

  const formattedDate: string = `${formatWithZero(currentDate.getDate())}-${
    formatWithZero(currentDate.getMonth() + 1)
  }-${currentDate.getFullYear()} ${formatWithZero(currentDate.getHours())}:${
    formatWithZero(currentDate.getMinutes())}:${formatWithZero(currentDate.getSeconds())}`;
  
  return (
    <p style={{ fontSize: "1rem", margin: 0 }} className={styles.titleDate}>
      Date: {formattedDate}
    </p>
  );
};

export default DateTime;

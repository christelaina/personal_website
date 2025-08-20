import React, { useState } from "react";
import Navbar from "../navbar";
import Footbar from "../footbar";
import "../../App.css";

const SECRET_PASSWORD = "ilovebugs"; // Change this to your desired password

const Secret = () => {
  const [password, setPassword] = useState("");
  const [unlocked, setUnlocked] = useState(false);
  const [error, setError] = useState("");

  const handleSubmit = (e) => {
    e.preventDefault();
    if (password === SECRET_PASSWORD) {
      setUnlocked(true);
      setError("");
    } else {
      setError("Incorrect password. Try again.");
    }
  };

  return (
    <div className="center">
      <div className="parent-content">
        <Navbar />
        <div className="content" style={{ flexDirection: 'column', alignItems: 'center', display: 'flex' }}>
          <h1 style={{ marginBottom: "1.5em", textAlign: "center", width: '100%' }}>secret</h1>
          {!unlocked ? (
            <form onSubmit={handleSubmit} style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "1em", width: "100%" }}>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="enter password"
                style={{
                  padding: "0.5em 1em",
                  fontSize: "1em",
                  width: "260px",
                  borderRadius: "8px",
                  border: "1.5px solid #b4bfa5",
                  background: "#fef9f5",
                  color: "#283106",
                  fontFamily: 'Jacques Francois',
                  marginBottom: "0.5em"
                }}
              />
              <button
                type="submit"
                style={{
                  padding: "0.5em 2em",
                  fontSize: "1em",
                  background: "#b4bfa5",
                  color: "#283106",
                  border: "1.5px solid #283106",
                  borderRadius: "18px",
                  cursor: "pointer",
                  fontFamily: 'Jacques Francois',
                  transition: "background 0.2s, color 0.2s, border 0.2s"
                }}
                onMouseOver={e => {
                  e.currentTarget.style.background = '#283106';
                  e.currentTarget.style.color = '#fef9f5';
                  e.currentTarget.style.border = '1.5px solid #b4bfa5';
                }}
                onMouseOut={e => {
                  e.currentTarget.style.background = '#b4bfa5';
                  e.currentTarget.style.color = '#283106';
                  e.currentTarget.style.border = '1.5px solid #283106';
                }}
              >
                unlock
              </button>
              {error && <div style={{ color: "red", marginTop: "0.5em" }}>{error}</div>}
            </form>
          ) : (
            <>
              <p style={{ marginBottom: "1em", textAlign: 'center', width: '100%' }}>correct! download:</p>
              <a
                href={require("../../assets/Original.txt")}
                download
                style={{
                  display: "inline-block",
                  padding: "0.5em 2em",
                  fontSize: "1em",
                  background: "#b4bfa5",
                  color: "#283106",
                  border: "1.5px solid #283106",
                  borderRadius: "18px",
                  cursor: "pointer",
                  fontFamily: 'Jacques Francois',
                  textDecoration: "none",
                  marginTop: "0",
                  transition: "background 0.2s, color 0.2s, border 0.2s"
                }}
                onMouseOver={e => {
                  e.currentTarget.style.background = '#283106';
                  e.currentTarget.style.color = '#fef9f5';
                  e.currentTarget.style.border = '1.5px solid #b4bfa5';
                }}
                onMouseOut={e => {
                  e.currentTarget.style.background = '#b4bfa5';
                  e.currentTarget.style.color = '#283106';
                  e.currentTarget.style.border = '1.5px solid #283106';
                }}
              >
                download file
              </a>
              <a
                href={require("../../assets/NameCheckVBA.txt")}
                download
                style={{
                  display: "inline-block",
                  padding: "0.5em 2em",
                  fontSize: "1em",
                  background: "#b4bfa5",
                  color: "#283106",
                  border: "1.5px solid #283106",
                  borderRadius: "18px",
                  cursor: "pointer",
                  fontFamily: 'Jacques Francois',
                  textDecoration: "none",
                  marginTop: "0",
                  transition: "background 0.2s, color 0.2s, border 0.2s"
                }}
                onMouseOver={e => {
                  e.currentTarget.style.background = '#283106';
                  e.currentTarget.style.color = '#fef9f5';
                  e.currentTarget.style.border = '1.5px solid #b4bfa5';
                }}
                onMouseOut={e => {
                  e.currentTarget.style.background = '#b4bfa5';
                  e.currentTarget.style.color = '#283106';
                  e.currentTarget.style.border = '1.5px solid #283106';
                }}
              >
                download file
              </a>
            </>
          )}
        </div>
        <Footbar />
      </div>
    </div>
  );
};

export default Secret;

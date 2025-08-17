import React from "react";
import { BrowserRouter as Router, Routes, Route } from "react-router-dom";
import Home from "./components/pages/home";
import Projects from "./components/pages/projects";
import About from "./components/pages/about";
import Secret from "./components/pages/secret";
import "./App.css";

const App = () => {
  return (
    <Router basename={process.env.PUBLIC_URL}>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/projects" element={<Projects />} />
        <Route path="/about" element={<About />} />
        <Route path="/secret" element={<Secret />} />
      </Routes>
    </Router>
  );
};

export default App;

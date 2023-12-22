import React from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom'
import Home from './components/pages/home'
import Projects from './components/pages/projects'
import About from './components/pages/about'
import Contact from './components/pages/contact'
import "./App.css"

const App = () => {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Home/>} />
        <Route path="/projects" element={<Projects/>} />
        <Route path="/about" element={<About/>} />
        <Route path="/contact" element={<Contact/>} />
      </Routes>        
    </Router>    
  )
}

export default App;

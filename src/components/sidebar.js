import { NavLink } from "react-router-dom";
import "./sidebar.css";


const Sidebar = () => {
  return (
    <nav className="sidebar">
        <div className="sidebar-elements">
          <ul>
            <li>
              <NavLink to="/">home</NavLink>
            </li>
            <li>
              <NavLink to="/about">about</NavLink>
            </li>
            <li>
              <NavLink to="/projects">projects</NavLink>
            </li>
            <li>
              <NavLink to="/contact">contact</NavLink>
            </li>
          </ul>
        </div>
    </nav>
  );
};

export default Sidebar;

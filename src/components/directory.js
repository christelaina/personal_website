import { NavLink } from "react-router-dom";
import "./directory.css";

const Directory = () => {
  return (
    <nav className="directory">
      <div className="directory-elements">
        <ul>
          <li>
            <NavLink to="/">home /</NavLink>
          </li>
          <li>
            <NavLink to="/about">about /</NavLink>
          </li>
          <li>
            <NavLink to="/projects">projects /</NavLink>
          </li>
        </ul>
      </div>
    </nav>
  );
};

export default Directory;

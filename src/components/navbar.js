import Directory from "./directory";
import Nameplate from "./nameplate";
import "../App.css";

const Navbar = () => {
  return (
    <>
      <div className="navbar">
        <Directory />
        <Nameplate />
      </div>
    </>
  );
};

export default Navbar;

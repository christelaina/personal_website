import "../App.css";

const Footbar = () => {
  return (
    <>
      <div className="footbar">
        <p>Toronto</p>
        <p>
          <a
            href="https://www.linkedin.com/in/elaina-araullo/"
            target="_blank"
            rel="noreferrer"
          >
            LinkedIn
          </a>{" "}
          ||{" "}
          <a
            href="https://github.com/christelaina"
            target="_blank"
            rel="noreferrer"
          >
            GitHub
          </a>{" "}
          ||{" "}
          <a href="../assets/Resume.pdf" target="_blank" rel="noreferrer">
            Resume
          </a>{" "}
          ||{" "}
          <a
            href="mailto:gea.christensen@gmail.com"
            target="_blank"
            rel="noreferrer"
          >
            Email
          </a>
        </p>
      </div>
    </>
  );
};

export default Footbar;

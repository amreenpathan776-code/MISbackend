module.exports = {
  apps: [
    {
      name: "server",
      script: "server.js",
      node_args: "--max_old_space_size=1024"
    }
  ]
}

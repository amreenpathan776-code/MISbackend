// tunnel.js
const tunnel = require('tunnel-ssh');
const fs = require('fs');

const config = {
  username: 'AdministratorDev',          // SSH user on UTI server
  host: '40.80.79.26',                   // UTI server IP
  port: 22,
  dstHost: '127.0.0.1',                  // destination inside UTI server
  dstPort: 3000,                         // UTI API port
  localHost: '127.0.0.1',                // EC2 local endpoint
  localPort: 11000,                      // local port for EC2 app to use
  privateKey: fs.readFileSync('/home/ec2-user/MISbackend/MISKeys/MISDashboard-key.pem'),
};

// ✅ Correct usage: call createTunnel(), not tunnel() directly
tunnel.createTunnel(config)
  .then(() => {
    console.log('✅ SSH Tunnel established: localhost:11000 → 40.80.79.26:3000');
  })
  .catch((error) => {
    console.error('❌ SSH Tunnel error:', error);
  });

import { Hono } from 'hono'
import { poweredBy } from 'hono/powered-by'

const app = new Hono();

app.use('*', poweredBy())

app.get('/hello', (c) => {
  return c.json({
    message: 'Hello from Hono!',
  })
})

export default app;

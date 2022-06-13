import { config } from 'dotenv'
// import { ApolloServer } from 'apollo-server-express'

import app from './app'
// import typeDefs from './typeDefs'
// import resolvers from './resolvers'
// import initMongoose from "./database"
import { routeNotFound } from './controllers/api.controller'

(async () => {
  config()
  // Initialize Mongoose
  // initMongoose()
  // Initialize Apollo Server
/*   const apolloServer = new ApolloServer({
    typeDefs,
    resolvers,
    context: ({ req }) => {
      console.log(req)
    },
  })
  await apolloServer.start()
  apolloServer.applyMiddleware({ app }) */
  app.use('*', routeNotFound)
  const port = process.env.PORT
  app.listen({ port }, () => console.log('Server listen http on port', port))
})()

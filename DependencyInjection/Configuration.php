<?php

namespace EXS\GoogleSheetsBundle\DependencyInjection;

use Symfony\Component\Config\Definition\Builder\TreeBuilder;
use Symfony\Component\Config\Definition\ConfigurationInterface;

/**
 * This is the class that validates and merges configuration from your app/config files.
 *
 * To learn more see {@link http://symfony.com/doc/current/cookbook/bundles/configuration.html}
 */
class Configuration implements ConfigurationInterface
{
    /**
     * {@inheritdoc}
     */
    public function getConfigTreeBuilder()
    {
        $treeBuilder = new TreeBuilder();
        $rootNode = $treeBuilder->root('exs_google_sheets');

        $rootNode
            ->children()
                ->scalarNode('application_name')
                    ->isRequired()
                    ->cannotBeEmpty()
                ->end()
                ->scalarNode('credentials')
                    ->isRequired()
                    ->cannotBeEmpty()
                ->end()
                ->scalarNode('client_secret')
                    ->isRequired()
                    ->cannotBeEmpty()
                ->end()            
            ->end()
        ;
        return $treeBuilder;
    }
}